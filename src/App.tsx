import React from 'react';
import { SharePointAuthProvider } from './SharePointAuthContext';
import SharePointLogin from './SharePointLogin';

const App: React.FC = () => {
  return (
    <SharePointAuthProvider>
      <div>
        <h1>My App</h1>
        <SharePointLogin />
        {/* Other components can also access the SharePointAuthContext */}
      </div>
    </SharePointAuthProvider>
  );
};

export default App;
