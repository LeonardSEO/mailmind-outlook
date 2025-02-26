import * as React from 'react';
import * as ReactDOM from 'react-dom/client';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import App from './components/App';

import '@fluentui/react/dist/css/fabric.min.css';

initializeIcons();

let isOfficeInitialized = false;

const title = 'MailMind';

const container = document.getElementById('container');
const root = ReactDOM.createRoot(container!);

const render = (Component: typeof App) => {
    root.render(
        <React.StrictMode>
            <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </React.StrictMode>
    );
};

/* Initial render showing a progress bar */
render(App);

/* Initialize Office */
Office.onReady(() => {
    isOfficeInitialized = true;
    render(App);
});