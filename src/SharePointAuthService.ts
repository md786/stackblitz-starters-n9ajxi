import axios from 'axios';

class SharePointAuthService {
  public authCode: string | null = null;
  public siteDetails: any = null;
  // public accessToken: string | null = null;

  public openLoginWindow(): void {
    const tenantId = 'ebd31e07-e19a-41fb-aebf-7d2ab2391202';
    const clientId = '7e489452-1f59-4673-bad9-4fde66e11bce';
    const redirectUri = encodeURIComponent(
      'https://stackblitz-starters-n9ajxi.stackblitz.io/'
    );

    const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?prompt=select_account&client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&scope=offline_access%20user.read%20sites.read.all`;
    const windowFeatures =
      'location=yes,height=600,width=800,scrollbars=yes,status=yes';
    const loginWindow = window.open(authUrl, '_blank', windowFeatures);

    // Polling to capture the authorization code
    const pollTimer = setInterval(async () => {
      try {
        const url = loginWindow.location.href;

        if (url.indexOf('code=') !== -1) {
          clearInterval(pollTimer);

          const urlParams = new URLSearchParams(loginWindow.location.search);
          const authorizationCode = urlParams.get('code');

          loginWindow.close();

          // Save the authorization code to the state
          this.authCode = authorizationCode;
          const accessToken = await this.getAccessToken(this.authCode);
          console.log(`Authorization code received: ${authorizationCode}`);
        }
      } catch (error) {
        // Continue polling or handle errors
      }
    }, 100);
  }

  public async getAccessToken(
    authorizationCode: string
  ): Promise<string | null> {
    const tenantId = 'ebd31e07-e19a-41fb-aebf-7d2ab2391202';
    const clientId = '7e489452-1f59-4673-bad9-4fde66e11bce';
    const redirectUri = encodeURIComponent(
      'https://stackblitz-starters-n9ajxi.stackblitz.io/'
    );
    const tokenEndpoint =
      'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token';

    const clientSecret = '-~68Q~.h9BgbYGOKVFIUlYbnZGYmQoOHpdHNsb9-'; // Use your actual client secret

    try {
      const response = await axios.post(tokenEndpoint, null, {
        params: {
          client_id: clientId,
          scope: 'offline_access user.read sites.read.all',
          code: authorizationCode,
          redirect_uri: redirectUri,
          grant_type: 'authorization_code',
          client_secret: clientSecret,
        },
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
      });

      return response.data.access_token;
    } catch (error) {
      console.error('Failed to obtain access token:', error);
      return null;
    }
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
