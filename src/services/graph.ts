import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";

const CLIENT_CREDENTIAL_SCOPE = "https://graph.microsoft.com/.default";

export interface AuthStatus {
  isAuthenticated: boolean;
  tenantId?: string | undefined;
  clientId?: string | undefined;
}

export class GraphService {
  private static instance: GraphService;
  private client: Client | undefined;
  private isInitialized = false;
  private msalApp: ConfidentialClientApplication | undefined;

  static getInstance(): GraphService {
    if (!GraphService.instance) {
      GraphService.instance = new GraphService();
    }
    return GraphService.instance;
  }

  private async initializeClient(): Promise<void> {
    if (this.isInitialized) return;

    try {
      // Priority 1: AUTH_TOKEN environment variable (direct token injection)
      const envToken = process.env.AUTH_TOKEN;
      if (envToken) {
        const validatedToken = this.validateToken(envToken);
        if (validatedToken) {
          this.client = Client.initWithMiddleware({
            authProvider: {
              getAccessToken: async () => validatedToken,
            },
          });
          this.isInitialized = true;
        }
        return;
      }

      // Priority 2: Client credentials via MSAL ConfidentialClientApplication
      const tenantId = process.env.AZURE_TENANT_ID;
      const clientId = process.env.AZURE_CLIENT_ID;
      const clientSecret = process.env.AZURE_CLIENT_SECRET;

      if (!tenantId || !clientId || !clientSecret) {
        console.error(
          "Missing required environment variables: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET"
        );
        return;
      }

      this.msalApp = new ConfidentialClientApplication({
        auth: {
          clientId,
          clientSecret,
          authority: `https://login.microsoftonline.com/${tenantId}`,
        },
      });

      // Verify we can acquire a token
      const result = await this.msalApp.acquireTokenByClientCredential({
        scopes: [CLIENT_CREDENTIAL_SCOPE],
      });

      if (!result) {
        return;
      }

      // Create Graph client with MSAL-backed auth provider for automatic token refresh
      this.client = Client.initWithMiddleware({
        authProvider: {
          getAccessToken: () => this.acquireToken(),
        },
      });

      this.isInitialized = true;
    } catch (error) {
      console.error("Failed to initialize Graph client:", error);
    }
  }

  private async acquireToken(): Promise<string> {
    if (!this.msalApp) {
      throw new Error("MSAL not initialized");
    }

    const result = await this.msalApp.acquireTokenByClientCredential({
      scopes: [CLIENT_CREDENTIAL_SCOPE],
    });

    if (!result) {
      throw new Error(
        "Failed to acquire access token. Check AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET."
      );
    }

    return result.accessToken;
  }

  async getAuthStatus(): Promise<AuthStatus> {
    await this.initializeClient();

    if (!this.client) {
      return { isAuthenticated: false };
    }

    try {
      await this.client.api("/organization").get();
      return {
        isAuthenticated: true,
        tenantId: process.env.AZURE_TENANT_ID,
        clientId: process.env.AZURE_CLIENT_ID,
      };
    } catch (error) {
      console.error("Error verifying auth:", error);
      return { isAuthenticated: false };
    }
  }

  async getClient(): Promise<Client> {
    await this.initializeClient();

    if (!this.client) {
      throw new Error(
        "Not authenticated. Check AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables."
      );
    }
    return this.client;
  }

  isAuthenticated(): boolean {
    return !!this.client && this.isInitialized;
  }

  validateToken(token: string): string | undefined {
    const tokenSplits = token.split(".");
    if (tokenSplits.length !== 3) {
      console.error("Invalid JWT token: missing claims");
      return undefined;
    }

    try {
      const payload = JSON.parse(atob(tokenSplits[1]));
      const audiences = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
      if (!audiences.includes("https://graph.microsoft.com")) {
        console.error("Invalid JWT token: Not a valid Microsoft Graph token");
        return undefined;
      }
    } catch (error) {
      console.error("Invalid JWT token: Failed to parse payload", error);
      return undefined;
    }

    return token;
  }
}
