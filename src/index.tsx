import * as React from 'react';
import * as ReactDOM from 'react-dom/client';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import App from './components/App';
import registerServiceWorker from './registerServiceWorker';

import './styles.less';
import '@fluentui/react/dist/css/fabric.min.css';

initializeIcons();

let isOfficeInitialized = false;

const title = 'outlook-addin-using-react-demo';

const container = document.getElementById('container');
const root = ReactDOM.createRoot(container!);

const render = (Component: typeof App) => {
    root.render(
        <React.StrictMode>
            <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </React.StrictMode>
    );
};

/* Render application after Office initializes */
Office.initialize = () => {
    isOfficeInitialized = true;
    render(App);
};

/* Initial render showing a progress bar */
render(App);

registerServiceWorker();