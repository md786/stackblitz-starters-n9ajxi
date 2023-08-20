import axios from 'axios';

class SharePointAuthService {
  public authCode: string | null = null;
  public siteDetails: any = null;
  // public accessToken: string | null = null;

  public openLoginWindow(): void {
    const tenantId = 'ebd31e07-e19a-41fb-aebf-7d2ab2391202';
    const clientId = '7e489452-1f59-4673-bad9-4fde66e11bce';
    const redirectUri = encodeURIComponent(
      'https://stackblitz-starters-bhy9pa.stackblitz.io/'
    );

    const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&scope=offline_access%20user.read%20sites.read.all`;
    const windowFeatures =
      'location=yes,height=600,width=800,scrollbars=yes,status=yes';
    const loginWindow = window.open(authUrl, '_blank', windowFeatures);

    // Polling to capture the authorization code
    const pollTimer = setInterval(() => {
      try {
        const url = loginWindow.location.href;

        if (url.indexOf('code=') !== -1) {
          clearInterval(pollTimer);

          const urlParams = new URLSearchParams(loginWindow.location.search);
          const authorizationCode = urlParams.get('code');

          loginWindow.close();

          // Save the authorization code to the state
          this.authCode = authorizationCode;

          console.log(`Authorization code received: ${authorizationCode}`);
        }
      } catch (error) {
        // Continue polling or handle errors
      }
    }, 100);
  }
  /*
  setAccessToken(token: string): void {
    this.accessToken = token;
  }*/
  public async fetchSiteDetails(accessToken: string): Promise<void> {
    const endpoint = 'https://graph.microsoft.com/v1.0/sites/root';
    try {
      const response = await axios.get(endpoint, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });
      this.siteDetails = response.data;
    } catch (error) {
      console.error('Failed to fetch site details:', error);
    }
  }
}

export default SharePointAuthService;
