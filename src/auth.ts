import { InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-browser'
import type { AccountInfo } from '@azure/msal-browser'
import { MSAL_AUTHORITY, MSAL_CLIENT_ID, SCOPES } from './config'

let pca: PublicClientApplication | null = null

export function isConfigured(): boolean {
  return MSAL_CLIENT_ID !== 'PENDIENTE-CLIENT-ID'
}

async function getMsal(): Promise<PublicClientApplication> {
  if (!pca) {
    pca = new PublicClientApplication({
      auth: {
        clientId: MSAL_CLIENT_ID,
        authority: MSAL_AUTHORITY,
        // La propia URL de la app (GitHub Pages), registrada como redirect URI
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: { cacheLocation: 'localStorage' },
    })
    await pca.initialize()
    const accounts = pca.getAllAccounts()
    if (accounts.length > 0) pca.setActiveAccount(accounts[0])
  }
  return pca
}

export async function getAccount(): Promise<AccountInfo | null> {
  if (!isConfigured()) return null
  const instance = await getMsal()
  return instance.getActiveAccount() ?? instance.getAllAccounts()[0] ?? null
}

export async function signIn(): Promise<AccountInfo> {
  const instance = await getMsal()
  const result = await instance.loginPopup({ scopes: SCOPES.base })
  if (result.account) instance.setActiveAccount(result.account)
  return result.account
}

export async function signOut(): Promise<void> {
  const instance = await getMsal()
  const account = await getAccount()
  await instance.logoutPopup(account ? { account } : undefined)
}

/** Token para Microsoft Graph; pide consentimiento solo si hace falta. */
export async function getToken(scopes: string[]): Promise<string> {
  const instance = await getMsal()
  const account = (await getAccount()) ?? (await signIn())
  try {
    const silent = await instance.acquireTokenSilent({ scopes, account })
    return silent.accessToken
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const interactive = await instance.acquireTokenPopup({ scopes })
      return interactive.accessToken
    }
    throw error
  }
}
