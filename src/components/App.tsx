/* global Office */
import * as React from 'react';
import {
    Stack,
    IStackTokens,
    TextField,
    PrimaryButton,
    IconButton,
    MessageBar,
    MessageBarType,
    Spinner,
    SpinnerSize,
    Text,
    Panel,
    initializeIcons
} from '@fluentui/react';
import OpenAI from 'openai';

interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

interface IMessage {
    role: 'user' | 'assistant';
    content: string;
}

interface IAppState {
    messages: IMessage[];
    currentMessage: string;
    isLoading: boolean;
    error: string | null;
    emailContext: {
        subject: string;
        body: string;
        sender: string;
        recipients: string[];
    } | null;
    isPanelOpen: boolean;
    apiKey: string;
    model: string;
    language: string;
}

const stackTokens: IStackTokens = {
    childrenGap: 10
};

class App extends React.Component<AppProps, IAppState> {
    private messagesEndRef = React.createRef<HTMLDivElement>();
    officeMailBoxItem = Office.context.mailbox.item;
    openai: OpenAI | null = null;

    constructor(props: AppProps) {
        super(props);
        this.state = {
            messages: [],
            currentMessage: '',
            isLoading: false,
            error: null,
            emailContext: null,
            isPanelOpen: false,
            apiKey: localStorage.getItem('openrouterApiKey') || '',
            model: localStorage.getItem('model') || '',
            language: localStorage.getItem('language') || 'English'
        };

        initializeIcons();
        this.initializeOpenAI();
        this.loadEmailContext();
    }

    scrollToBottom = () => {
        this.messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
    };

    componentDidUpdate(_prevProps: AppProps, prevState: IAppState) {
        if (prevState.messages.length !== this.state.messages.length) {
            this.scrollToBottom();
        }
    }

    initializeOpenAI = () => {
        if (this.state.apiKey) {
            this.openai = new OpenAI({
                baseURL: 'https://openrouter.ai/api/v1',
                apiKey: this.state.apiKey,
                defaultHeaders: {
                    'HTTP-Referer': window.location.origin,
                    'X-Title': 'MailMind',
                },
                dangerouslyAllowBrowser: true
            });
        }
    };

    loadEmailContext = async () => {
        try {
            this.setState({ isLoading: true });
            
            const subject = this.officeMailBoxItem.subject || '';
            
            let body = '';
            await new Promise((resolve) => {
                this.officeMailBoxItem.body.getAsync(Office.CoercionType.Text, (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        body = result.value;
                    }
                    resolve(null);
                });
            });

            const sender = this.officeMailBoxItem.from ? this.officeMailBoxItem.from.emailAddress : '';
            const recipients = this.officeMailBoxItem.to ? 
                this.officeMailBoxItem.to.map(r => r.emailAddress) : 
                [];

            this.setState({
                emailContext: {
                    subject,
                    body,
                    sender,
                    recipients
                },
                isLoading: false
            });

            this.setState(prevState => ({
                messages: [
                    ...prevState.messages,
                    {
                        role: 'assistant',
                        content: 'Hi! I\'m your AI email assistant. I can help you with:' +
                            '\n• Writing and improving emails' +
                            '\n• Summarizing email threads' +
                            '\n• Extracting key information' +
                            '\n• Translating content' +
                            '\n• Suggesting responses' +
                            '\nWhat would you like me to help you with?' +
                            '\n\nTip: Click the settings icon to configure your OpenRouter API key and model.'
                    }
                ]
            }));

        } catch (error) {
            this.setState({
                error: 'Failed to load email context',
                isLoading: false
            });
        }
    };

    handleMessageChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        this.setState({ currentMessage: newValue || '' });
    };

    handleReply = async (messageContent: string) => {
        try {
            // Get the current item
            const item = Office.context.mailbox.item;
            
            if (!item) {
                this.setState({ error: 'No active email compose window found' });
                return;
            }

            // Extract the reply content between --- markers
            const matches = messageContent.match(/---\n([\s\S]*?)\n---/);
            if (!matches || !matches[1]) {
                this.setState({ error: 'No properly formatted reply found. Reply should be between --- markers.' });
                return;
            }

            const replyContent = matches[1].trim();

            // Create a new message with the reply, preserving line breaks
            Office.context.mailbox.displayNewMessageForm({
                toRecipients: item.to?.map(recipient => recipient.emailAddress) || [],
                subject: `Re: ${item.subject || ''}`,
                htmlBody: replyContent.replace(/\n/g, '<br>')
            });

        } catch (error) {
            this.setState({
                error: 'Failed to insert reply: ' + (error instanceof Error ? error.message : 'Unknown error')
            });
        }
    };

    handleSettingsSave = () => {
        const { apiKey, model, language } = this.state;
        localStorage.setItem('openrouterApiKey', apiKey);
        localStorage.setItem('model', model);
        localStorage.setItem('language', language);
        this.initializeOpenAI();
        this.setState({ isPanelOpen: false });
    };

    handleSendMessage = async (messageContent: string) => {
        const { emailContext, model, language } = this.state;
        
        if (!messageContent.trim() || !this.openai || !model) {
            this.setState({ error: 'Please configure your API key and model in settings first.' });
            return;
        }

        this.setState(prevState => ({
            messages: [
                ...prevState.messages,
                { role: 'user', content: messageContent }
            ],
            currentMessage: '',
            isLoading: true,
            error: null
        }));

        try {
            const systemPrompt = `You are an expert email assistant. Your task is to help compose professional and effective email responses in ${language}.

Current email context:
Subject: ${emailContext?.subject}
From: ${emailContext?.sender}
To: ${emailContext?.recipients.join(', ')}

Email Body:
${emailContext?.body}

Guidelines:
- Keep responses concise and professional
- Maintain appropriate tone based on context
- Format response in plain text suitable for email
- Focus on addressing the key points
- Be direct but polite
- Always respond in ${language}

Please provide assistance based on this context and the user's request.`;

            const completion = await this.openai.chat.completions.create({
                model: model,
                messages: [
                    { role: 'system', content: systemPrompt },
                    ...this.state.messages.map(m => ({ role: m.role, content: m.content })),
                    { role: 'user', content: messageContent }
                ],
                temperature: 0.2,
                top_p: 0.9,
                max_tokens: 300,
                frequency_penalty: 0.5,
                presence_penalty: 0.5
            });

            const response = completion.choices[0].message.content;

            this.setState(prevState => ({
                messages: [
                    ...prevState.messages,
                    { role: 'assistant', content: response || 'Sorry, I couldn\'t generate a response.' }
                ],
                isLoading: false
            }));

        } catch (error) {
            this.setState({
                error: 'Failed to get AI response. Please check your API key and model name.',
                isLoading: false
            });
        }
    };

    handleRegenerateResponse = () => {
        const lastUserMessage = this.state.messages.filter(m => m.role === 'user').pop();
        if (lastUserMessage) {
            this.handleSendMessage(lastUserMessage.content);
        }
    };

    render() {
        const { 
            messages, 
            currentMessage, 
            isLoading, 
            error, 
            emailContext,
            isPanelOpen,
            apiKey,
            model,
            language
        } = this.state;

        if (!this.props.isOfficeInitialized) {
            return (
                <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { height: '100vh' } }}>
                    <Spinner size={SpinnerSize.large} label="Loading Office.js..." />
                </Stack>
            );
        }

        return (
            <Stack 
                tokens={stackTokens}
                styles={{
                    root: {
                        padding: '16px 20px',
                        height: '100vh',
                        backgroundColor: '#ffffff'
                    }
                }}
            >
                {/* Header */}
                <Stack.Item>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { marginBottom: 8 } }}>
                        <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                            <Text variant="large" styles={{ root: { fontWeight: 600, color: '#616161' } }}>Current Email</Text>
                        </Stack>
                        <IconButton
                            iconProps={{ iconName: 'Settings' }}
                            title="Settings"
                            ariaLabel="Settings"
                            styles={{
                                root: {
                                    color: '#616161'
                                }
                            }}
                            onClick={() => this.setState({ isPanelOpen: true })}
                        />
                    </Stack>
                </Stack.Item>

                {/* Settings Panel */}
                <Panel
                    isOpen={isPanelOpen}
                    onDismiss={() => this.setState({ isPanelOpen: false })}
                    headerText="Email Assistant Settings"
                    closeButtonAriaLabel="Close"
                    styles={{
                        main: {
                            boxShadow: '0 8px 32px rgba(0,0,0,0.12)'
                        }
                    }}
                >
                    <Stack tokens={{ childrenGap: 20 }}>
                        <TextField
                            label="OpenRouter API Key"
                            value={apiKey}
                            onChange={(_ev, newValue) => this.setState({ apiKey: newValue || '' })}
                            type="password"
                            styles={{
                                fieldGroup: {
                                    borderRadius: 4
                                }
                            }}
                        />
                        <TextField
                            label="Model Name"
                            value={model}
                            onChange={(_ev, newValue) => this.setState({ model: newValue || '' })}
                            placeholder="e.g., openai/gpt-4, anthropic/claude-3-opus"
                            styles={{
                                fieldGroup: {
                                    borderRadius: 4
                                }
                            }}
                        />
                        <TextField
                            label="Preferred Language"
                            value={language}
                            onChange={(_ev, newValue) => this.setState({ language: newValue || 'English' })}
                            placeholder="e.g., English, Dutch, French"
                            styles={{
                                fieldGroup: {
                                    borderRadius: 4
                                }
                            }}
                        />
                        <PrimaryButton 
                            text="Save Settings" 
                            onClick={this.handleSettingsSave}
                            styles={{
                                root: {
                                    borderRadius: 4,
                                    marginTop: 10
                                }
                            }}
                        />
                    </Stack>
                </Panel>

                {/* Messages Container */}
                <Stack.Item grow styles={{ root: { overflowY: 'auto', margin: '16px -20px', padding: '0 20px' } }}>
                    <Stack tokens={{ childrenGap: 16 }}>
                        {emailContext && (
                            <Stack.Item>
                                <Stack 
                                    tokens={{ childrenGap: 8 }} 
                                    styles={{ 
                                        root: { 
                                            padding: 16,
                                            backgroundColor: '#f8f9fa',
                                            borderRadius: 8,
                                            border: '1px solid #e9ecef'
                                        } 
                                    }}
                                >
                                    <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#495057' } }}>Current Email</Text>
                                    <Text styles={{ root: { color: '#495057' } }}>Subject: {emailContext.subject}</Text>
                                    <Text styles={{ root: { color: '#495057' } }}>From: {emailContext.sender}</Text>
                                    <Text styles={{ root: { color: '#495057' } }}>To: {emailContext.recipients.join(', ')}</Text>
                                </Stack>
                            </Stack.Item>
                        )}
                        {messages.map((msg, index) => (
                            <Stack.Item key={index}>
                                <Stack tokens={{ childrenGap: 8 }}>
                                    <Text 
                                        variant="mediumPlus" 
                                        styles={{ 
                                            root: { 
                                                fontWeight: 600,
                                                color: msg.role === 'user' ? '#1a73e8' : '#34a853'
                                            } 
                                        }}
                                    >
                                        {msg.role === 'user' ? 'You: ' : 'Assistant: '}
                                    </Text>
                                    <Stack 
                                        styles={{ 
                                            root: { 
                                                backgroundColor: msg.role === 'user' ? '#f8f9fa' : '#ffffff',
                                                padding: 16,
                                                borderRadius: 8,
                                                border: '1px solid #e9ecef'
                                            } 
                                        }}
                                    >
                                        <Text styles={{ root: { whiteSpace: 'pre-wrap', color: '#212529' } }}>
                                            {msg.content}
                                        </Text>
                                        {msg.role === 'assistant' && (
                                            <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 12 } }}>
                                                {msg.content.includes('---') && (
                                                    <PrimaryButton
                                                        text="Use as Reply"
                                                        onClick={() => this.handleReply(msg.content)}
                                                        styles={{
                                                            root: {
                                                                borderRadius: 4,
                                                                backgroundColor: '#34a853',
                                                                border: 'none'
                                                            }
                                                        }}
                                                    />
                                                )}
                                                {!msg.content.includes('Hi! I\'m your AI email assistant') && (
                                                    <IconButton
                                                        iconProps={{ iconName: 'Refresh' }}
                                                        title="Regenerate response"
                                                        ariaLabel="Regenerate response"
                                                        onClick={this.handleRegenerateResponse}
                                                        styles={{
                                                            root: {
                                                                color: '#616161',
                                                                marginLeft: 8
                                                            }
                                                        }}
                                                    />
                                                )}
                                            </Stack>
                                        )}
                                    </Stack>
                                </Stack>
                            </Stack.Item>
                        ))}
                        <div ref={this.messagesEndRef} />
                        {isLoading && (
                            <Stack.Item>
                                <Spinner size={SpinnerSize.small} styles={{ root: { padding: 20 } }} />
                            </Stack.Item>
                        )}
                        {error && (
                            <Stack.Item>
                                <MessageBar 
                                    messageBarType={MessageBarType.error}
                                    styles={{
                                        root: {
                                            borderRadius: 4
                                        }
                                    }}
                                >
                                    {error}
                                </MessageBar>
                            </Stack.Item>
                        )}
                    </Stack>
                </Stack.Item>

                {/* Input Area */}
                <Stack.Item>
                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <Stack.Item grow>
                            <TextField
                        multiline
                                rows={2}
                                value={currentMessage}
                                onChange={this.handleMessageChange}
                                placeholder="Type your message..."
                                styles={{
                                    fieldGroup: {
                                        borderRadius: 4
                                    }
                                }}
                            />
                        </Stack.Item>
                        <PrimaryButton
                            text="Send"
                            onClick={() => this.handleSendMessage(currentMessage)}
                            disabled={isLoading || !currentMessage.trim()}
                            styles={{
                                root: {
                                    borderRadius: 4,
                                    height: 'auto',
                                    alignSelf: 'flex-end'
                                }
                            }}
                        />
                    </Stack>
                </Stack.Item>
            </Stack>
        );
    }
}

export default App;