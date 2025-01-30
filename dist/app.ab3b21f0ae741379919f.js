/*! For license information please see app.ab3b21f0ae741379919f.js.LICENSE.txt */
"use strict";(self.webpackChunk=self.webpackChunk||[]).push([[524],{58923:(e,t,n)=>{n.r(t)},6043:function(e,t,n){var o,r=this&&this.__extends||(o=function(e,t){return o=Object.setPrototypeOf||{__proto__:[]}instanceof Array&&function(e,t){e.__proto__=t}||function(e,t){for(var n in t)Object.prototype.hasOwnProperty.call(t,n)&&(e[n]=t[n])},o(e,t)},function(e,t){if("function"!=typeof t&&null!==t)throw new TypeError("Class extends value "+String(t)+" is not a constructor or null");function n(){this.constructor=e}o(e,t),e.prototype=null===t?Object.create(t):(n.prototype=t.prototype,new n)}),a=this&&this.__awaiter||function(e,t,n,o){return new(n||(n=Promise))((function(r,a){function i(e){try{l(o.next(e))}catch(e){a(e)}}function s(e){try{l(o.throw(e))}catch(e){a(e)}}function l(e){var t;e.done?r(e.value):(t=e.value,t instanceof n?t:new n((function(e){e(t)}))).then(i,s)}l((o=o.apply(e,t||[])).next())}))},i=this&&this.__generator||function(e,t){var n,o,r,a={label:0,sent:function(){if(1&r[0])throw r[1];return r[1]},trys:[],ops:[]},i=Object.create(("function"==typeof Iterator?Iterator:Object).prototype);return i.next=s(0),i.throw=s(1),i.return=s(2),"function"==typeof Symbol&&(i[Symbol.iterator]=function(){return this}),i;function s(s){return function(l){return function(s){if(n)throw new TypeError("Generator is already executing.");for(;i&&(i=0,s[0]&&(a=0)),a;)try{if(n=1,o&&(r=2&s[0]?o.return:s[0]?o.throw||((r=o.return)&&r.call(o),0):o.next)&&!(r=r.call(o,s[1])).done)return r;switch(o=0,r&&(s=[2&s[0],r.value]),s[0]){case 0:case 1:r=s;break;case 4:return a.label++,{value:s[1],done:!1};case 5:a.label++,o=s[1],s=[0];continue;case 7:s=a.ops.pop(),a.trys.pop();continue;default:if(!((r=(r=a.trys).length>0&&r[r.length-1])||6!==s[0]&&2!==s[0])){a=0;continue}if(3===s[0]&&(!r||s[1]>r[0]&&s[1]<r[3])){a.label=s[1];break}if(6===s[0]&&a.label<r[1]){a.label=r[1],r=s;break}if(r&&a.label<r[2]){a.label=r[2],a.ops.push(s);break}r[2]&&a.ops.pop(),a.trys.pop();continue}s=t.call(e,a)}catch(e){s=[6,e],o=0}finally{n=r=0}if(5&s[0])throw s[1];return{value:s[0]?s[1]:void 0,done:!0}}([s,l])}}},s=this&&this.__spreadArray||function(e,t,n){if(n||2===arguments.length)for(var o,r=0,a=t.length;r<a;r++)!o&&r in t||(o||(o=Array.prototype.slice.call(t,0,r)),o[r]=t[r]);return e.concat(o||Array.prototype.slice.call(t))};Object.defineProperty(t,"__esModule",{value:!0});var l=n(63696),c=n(22788),u=n(31150),d={childrenGap:10},p=function(e){function t(t){var n=e.call(this,t)||this;return n.messagesEndRef=l.createRef(),n.officeMailBoxItem=Office.context.mailbox.item,n.openai=null,n.scrollToBottom=function(){var e;null===(e=n.messagesEndRef.current)||void 0===e||e.scrollIntoView({behavior:"smooth"})},n.initializeOpenAI=function(){n.state.apiKey&&(n.openai=new u.default({baseURL:"https://openrouter.ai/api/v1",apiKey:n.state.apiKey,defaultHeaders:{"HTTP-Referer":window.location.origin,"X-Title":"MailMind"},dangerouslyAllowBrowser:!0}))},n.loadEmailContext=function(){return a(n,void 0,void 0,(function(){var e,t,n,o,r=this;return i(this,(function(a){switch(a.label){case 0:return a.trys.push([0,2,,3]),this.setState({isLoading:!0}),e=this.officeMailBoxItem.subject||"",t="",[4,new Promise((function(e){r.officeMailBoxItem.body.getAsync(Office.CoercionType.Text,(function(n){n.status===Office.AsyncResultStatus.Succeeded&&(t=n.value),e(null)}))}))];case 1:return a.sent(),n=this.officeMailBoxItem.from?this.officeMailBoxItem.from.emailAddress:"",o=this.officeMailBoxItem.to?this.officeMailBoxItem.to.map((function(e){return e.emailAddress})):[],this.setState({emailContext:{subject:e,body:t,sender:n,recipients:o},isLoading:!1}),this.setState((function(e){return{messages:s(s([],e.messages,!0),[{role:"assistant",content:"Hi! I'm your AI email assistant. I can help you with:\n• Writing and improving emails\n• Summarizing email threads\n• Extracting key information\n• Translating content\n• Suggesting responses\nWhat would you like me to help you with?\n\nTip: Click the settings icon to configure your OpenRouter API key and model."}],!1)}})),[3,3];case 2:return a.sent(),this.setState({error:"Failed to load email context",isLoading:!1}),[3,3];case 3:return[2]}}))}))},n.handleMessageChange=function(e,t){n.setState({currentMessage:t||""})},n.handleReply=function(e){return a(n,void 0,void 0,(function(){var t,n,o,r;return i(this,(function(a){try{if(!(t=Office.context.mailbox.item))return this.setState({error:"No active email compose window found"}),[2];if(!(n=e.match(/---\n([\s\S]*?)\n---/))||!n[1])return this.setState({error:"No properly formatted reply found. Reply should be between --- markers."}),[2];o=n[1].trim(),Office.context.mailbox.displayNewMessageForm({toRecipients:(null===(r=t.to)||void 0===r?void 0:r.map((function(e){return e.emailAddress})))||[],subject:"Re: ".concat(t.subject||""),htmlBody:o.replace(/\n/g,"<br>")})}catch(e){this.setState({error:"Failed to insert reply: "+(e instanceof Error?e.message:"Unknown error")})}return[2]}))}))},n.handleSettingsSave=function(){var e=n.state,t=e.apiKey,o=e.model,r=e.language;localStorage.setItem("openrouterApiKey",t),localStorage.setItem("model",o),localStorage.setItem("language",r),n.initializeOpenAI(),n.setState({isPanelOpen:!1})},n.handleSendMessage=function(){return a(n,void 0,void 0,(function(){var e,t,n,o,r,a,l,c;return i(this,(function(i){switch(i.label){case 0:if(e=this.state,t=e.currentMessage,n=e.emailContext,o=e.model,r=e.language,!t.trim()||!this.openai||!o)return this.setState({error:"Please configure your API key and model in settings first."}),[2];this.setState((function(e){return{messages:s(s([],e.messages,!0),[{role:"user",content:t}],!1),currentMessage:"",isLoading:!0,error:null}})),i.label=1;case 1:return i.trys.push([1,3,,4]),a="You are an expert email assistant. Your task is to help compose professional and effective email responses in ".concat(r,".\n\nCurrent email context:\nSubject: ").concat(null==n?void 0:n.subject,"\nFrom: ").concat(null==n?void 0:n.sender,"\nTo: ").concat(null==n?void 0:n.recipients.join(", "),"\n\nEmail Body:\n").concat(null==n?void 0:n.body,"\n\nGuidelines:\n- Keep responses concise and professional\n- Maintain appropriate tone based on context\n- Format response in plain text suitable for email\n- Focus on addressing the key points\n- Be direct but polite\n- Always respond in ").concat(r,"\n\nPlease provide assistance based on this context and the user's request."),[4,this.openai.chat.completions.create({model:o,messages:s(s([{role:"system",content:a}],this.state.messages.map((function(e){return{role:e.role,content:e.content}})),!0),[{role:"user",content:t}],!1),temperature:.2,top_p:.9,max_tokens:300,frequency_penalty:.5,presence_penalty:.5})];case 2:return l=i.sent(),c=l.choices[0].message.content,this.setState((function(e){return{messages:s(s([],e.messages,!0),[{role:"assistant",content:c||"Sorry, I couldn't generate a response."}],!1),isLoading:!1}})),[3,4];case 3:return i.sent(),this.setState({error:"Failed to get AI response. Please check your API key and model name.",isLoading:!1}),[3,4];case 4:return[2]}}))}))},n.state={messages:[],currentMessage:"",isLoading:!1,error:null,emailContext:null,isPanelOpen:!1,apiKey:localStorage.getItem("openrouterApiKey")||"",model:localStorage.getItem("model")||"",language:localStorage.getItem("language")||"English"},(0,c.initializeIcons)(),n.initializeOpenAI(),n.loadEmailContext(),n}return r(t,e),t.prototype.componentDidUpdate=function(e,t){t.messages.length!==this.state.messages.length&&this.scrollToBottom()},t.prototype.render=function(){var e=this,t=this.state,n=t.messages,o=t.currentMessage,r=t.isLoading,a=t.error,i=t.emailContext,s=t.isPanelOpen,u=t.apiKey,p=t.model,f=t.language;return this.props.isOfficeInitialized?l.createElement(c.Stack,{tokens:d,styles:{root:{padding:"16px 20px",height:"100vh",backgroundColor:"#ffffff"}}},l.createElement(c.Stack.Item,null,l.createElement(c.Stack,{horizontal:!0,horizontalAlign:"space-between",verticalAlign:"center",styles:{root:{marginBottom:8}}},l.createElement(c.Stack,{horizontal:!0,tokens:{childrenGap:12},verticalAlign:"center"},l.createElement(c.Text,{variant:"large",styles:{root:{fontWeight:600,color:"#616161"}}},"Current Email")),l.createElement(c.IconButton,{iconProps:{iconName:"Settings"},title:"Settings",ariaLabel:"Settings",styles:{root:{color:"#616161"}},onClick:function(){return e.setState({isPanelOpen:!0})}}))),l.createElement(c.Panel,{isOpen:s,onDismiss:function(){return e.setState({isPanelOpen:!1})},headerText:"Email Assistant Settings",closeButtonAriaLabel:"Close",styles:{main:{boxShadow:"0 8px 32px rgba(0,0,0,0.12)"}}},l.createElement(c.Stack,{tokens:{childrenGap:20}},l.createElement(c.TextField,{label:"OpenRouter API Key",value:u,onChange:function(t,n){return e.setState({apiKey:n||""})},type:"password",styles:{fieldGroup:{borderRadius:4}}}),l.createElement(c.TextField,{label:"Model Name",value:p,onChange:function(t,n){return e.setState({model:n||""})},placeholder:"e.g., openai/gpt-4, anthropic/claude-3-opus",styles:{fieldGroup:{borderRadius:4}}}),l.createElement(c.TextField,{label:"Preferred Language",value:f,onChange:function(t,n){return e.setState({language:n||"English"})},placeholder:"e.g., English, Dutch, French",styles:{fieldGroup:{borderRadius:4}}}),l.createElement(c.PrimaryButton,{text:"Save Settings",onClick:this.handleSettingsSave,styles:{root:{borderRadius:4,marginTop:10}}}))),l.createElement(c.Stack.Item,{grow:!0,styles:{root:{overflowY:"auto",margin:"16px -20px",padding:"0 20px"}}},l.createElement(c.Stack,{tokens:{childrenGap:16}},i&&l.createElement(c.Stack.Item,null,l.createElement(c.Stack,{tokens:{childrenGap:8},styles:{root:{padding:16,backgroundColor:"#f8f9fa",borderRadius:8,border:"1px solid #e9ecef"}}},l.createElement(c.Text,{variant:"mediumPlus",styles:{root:{fontWeight:600,color:"#495057"}}},"Current Email"),l.createElement(c.Text,{styles:{root:{color:"#495057"}}},"Subject: ",i.subject),l.createElement(c.Text,{styles:{root:{color:"#495057"}}},"From: ",i.sender),l.createElement(c.Text,{styles:{root:{color:"#495057"}}},"To: ",i.recipients.join(", ")))),n.map((function(t,n){return l.createElement(c.Stack.Item,{key:n},l.createElement(c.Stack,{tokens:{childrenGap:8}},l.createElement(c.Text,{variant:"mediumPlus",styles:{root:{fontWeight:600,color:"user"===t.role?"#1a73e8":"#34a853"}}},"user"===t.role?"You: ":"Assistant: "),l.createElement(c.Stack,{styles:{root:{backgroundColor:"user"===t.role?"#f8f9fa":"#ffffff",padding:16,borderRadius:8,border:"1px solid #e9ecef"}}},l.createElement(c.Text,{styles:{root:{whiteSpace:"pre-wrap",color:"#212529"}}},t.content),"assistant"===t.role&&l.createElement(c.PrimaryButton,{text:"Use as Reply",onClick:function(){return e.handleReply(t.content)},styles:{root:{marginTop:12,borderRadius:4,backgroundColor:"#34a853",border:"none"}}}))))})),l.createElement("div",{ref:this.messagesEndRef}),r&&l.createElement(c.Stack.Item,null,l.createElement(c.Spinner,{size:c.SpinnerSize.small,styles:{root:{padding:20}}})),a&&l.createElement(c.Stack.Item,null,l.createElement(c.MessageBar,{messageBarType:c.MessageBarType.error,styles:{root:{borderRadius:4}}},a)))),l.createElement(c.Stack.Item,null,l.createElement(c.Stack,{horizontal:!0,tokens:{childrenGap:8}},l.createElement(c.Stack.Item,{grow:!0},l.createElement(c.TextField,{multiline:!0,rows:2,value:o,onChange:this.handleMessageChange,placeholder:"Type your message...",styles:{fieldGroup:{borderRadius:4}}})),l.createElement(c.PrimaryButton,{text:"Send",onClick:this.handleSendMessage,disabled:r||!o.trim(),styles:{root:{borderRadius:4,height:"auto",alignSelf:"flex-end"}}})))):l.createElement(c.Stack,{horizontalAlign:"center",verticalAlign:"center",styles:{root:{height:"100vh"}}},l.createElement(c.Spinner,{size:c.SpinnerSize.large,label:"Loading Office.js..."}))},t}(l.Component);t.default=p},34111:(e,t,n)=>{var o=n(63696),r=n(7470),a=n(94134),i=n(6043),s=n(83523);n(58923),n(56569),(0,a.initializeIcons)();var l=!1,c=document.getElementById("container"),u=r.createRoot(c),d=function(e){u.render(o.createElement(o.StrictMode,null,o.createElement(e,{title:"outlook-addin-using-react-demo",isOfficeInitialized:l})))};Office.initialize=function(){l=!0,d(i.default)},d(i.default),(0,s.default)()},83523:(e,t,n)=>{n.r(t),n.d(t,{default:()=>r,unregister:()=>i});const o=Boolean("localhost"===window.location.hostname||"[::1]"===window.location.hostname||window.location.hostname.match(/^127(?:\.(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){3}$/));function r(){if("serviceWorker"in navigator){if(new URL(process.env.PUBLIC_URL,window.location).origin!==window.location.origin)return;window.addEventListener("load",(()=>{const e=`${process.env.PUBLIC_URL}/service-worker.js`;o?(function(e){fetch(e).then((t=>{404===t.status||-1===t.headers.get("content-type").indexOf("javascript")?navigator.serviceWorker.ready.then((e=>{e.unregister().then((()=>{window.location.reload()}))})):a(e)})).catch((()=>{console.log("No internet connection found. App is running in offline mode.")}))}(e),navigator.serviceWorker.ready.then((()=>{console.log("This web app is being served cache-first by a service worker. To learn more, visit https://goo.gl/SC7cgQ")}))):a(e)}))}}function a(e){navigator.serviceWorker.register(e).then((e=>{e.onupdatefound=()=>{const t=e.installing;t.onstatechange=()=>{"installed"===t.state&&(navigator.serviceWorker.controller?console.log("New content is available; please refresh."):console.log("Content is cached for offline use."))}}})).catch((e=>{console.error("Error during service worker registration:",e)}))}function i(){"serviceWorker"in navigator&&navigator.serviceWorker.ready.then((e=>{e.unregister()}))}}},e=>{e.O(0,[96],(()=>e(e.s=34111))),e.O()}]);
//# sourceMappingURL=app.ab3b21f0ae741379919f.js.map