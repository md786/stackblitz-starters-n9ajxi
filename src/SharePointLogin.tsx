import React, { useState } from 'react';
import { useSharePointAuth } from './SharePointAuthContext';

const SharePointLogin: React.FC = () => {
  const authService = useSharePointAuth();
  const [accessToken, setAccessToken] = useState<string | null>(null);

  const handleLogin = async () => {
    authService.openLoginWindow();
    // Assume we get access token somehow after successful login
    //setAccessToken(authService.authCode);

    // Fetch site details using the access token
    if (accessToken) {
      await authService.fetchSiteDetails(accessToken);
    }
  };

  return (
    <div>
      <button onClick={handleLogin}>Login to SharePoint</button>
      {/* Access authCode or other properties/methods from authService as needed */}
      {authService.authCode && (
        <p>Authorization Code: {authService.authCode}</p>
      )}
      {authService.siteDetails && (
        <pre>{JSON.stringify(authService.siteDetails, null, 2)}</pre>
      )}
    </div>
  );
};

export default SharePointLogin;
