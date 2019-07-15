import React from 'react';
import ReactDOM from 'react-dom';
import { App } from './components/App';
import { mergeStyles } from 'office-ui-fabric-react';

// Inject some global styles
mergeStyles({
  selectors: {
    ':global(body), :global(html), :global(#app)': {
      margin: 0,
      padding: 0,
      height: '100vh'
    }
  }
});

ReactDOM.render(<App 
  ref={(app_component) => {(window as any).react_app_component=app_component;}} // in order to call the app from outside react.
  />, 
  document.getElementById('app')
  );
