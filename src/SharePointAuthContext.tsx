import React, { createContext, useContext, useState, ReactNode } from 'react';
import SharePointAuthService from './SharePointAuthService';

interface SharePointAuthContextProps {
  children: ReactNode;
}

const SharePointAuthContext = createContext<SharePointAuthService | undefined>(
  undefined
);

export const useSharePointAuth = (): SharePointAuthService => {
  const context = useContext(SharePointAuthContext);
  if (!context) {
    throw new Error(
      'useSharePointAuth must be used within a SharePointAuthProvider'
    );
  }
  return context;
};

export const SharePointAuthProvider: React.FC<SharePointAuthContextProps> = ({
  children,
}) => {
  const [authService] = useState(new SharePointAuthService());

  return (
    <SharePointAuthContext.Provider value={authService}>
      {children}
    </SharePointAuthContext.Provider>
  );
};
