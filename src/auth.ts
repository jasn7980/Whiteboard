import { PublicClientApplication } from '@azure/msal-browser'
import type { AccountInfo } from '@azure/msal-browser'
import { ALL_SCOPES, MSAL_AUTHORITY, MSAL_CLIENT_ID } from './config'

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
      system: {
        // Más paciencia para la renovación silenciosa del token: en la
        // Chromebook el iframe oculto de Microsoft a veces tarda y MSAL
        // abortaba con "timed_out" a los pocos segundos.
        iframeHashTimeout: 15000,
      },
    })
    await pca.initialize()
    // Flujo por redirección: al volver de Microsoft, aquí se completa el
    // inicio de sesión y se limpia el estado de interacción.
    const redirectResult = await pca.handleRedirectPromise().catch(() => null)
    if (redirectResult?.account) pca.setActiveAccount(redirectResult.account)
    const accounts = pca.getAllAccounts()
    if (accounts.length > 0 && !pca.getActiveAccount()) {
      pca.setActiveAccount(accounts[0])
    }
  }
  return pca
}

export async function getAccount(): Promise<AccountInfo | null> {
  if (!isConfigured()) return null
  const instance = await getMsal()
  return instance.getActiveAccount() ?? instance.getAllAccounts()[0] ?? null
}

/** Traduce errores técnicos de MSAL a instrucciones accionables. */
function friendlyAuthError(e: unknown): Error {
  const msg = e instanceof Error ? e.message : String(e)
  if (msg.includes('interaction_in_progress')) {
    return new Error(
      'Hay un inicio de sesión atascado en esta pestaña. Ciérrala por completo, vuelve a abrir el tablero e inténtalo de nuevo.',
    )
  }
  if (msg.includes('timed_out') || msg.includes('monitor_window_timeout')) {
    return new Error('Microsoft tardó demasiado en responder. Inténtalo de nuevo.')
  }
  return e instanceof Error ? e : new Error(msg)
}

export async function signIn(): Promise<void> {
  const instance = await getMsal()
  // Redirección (no popup): los popups de login son frágiles en ChromeOS y
  // dejaban el estado "interaction_in_progress" atascado. La página va a
  // Microsoft y regresa; getMsal() completa el flujo con
  // handleRedirectPromise() al cargar.
  try {
    await instance.loginRedirect({ scopes: ALL_SCOPES })
  } catch (e) {
    throw friendlyAuthError(e)
  }
}

export async function signOut(): Promise<void> {
  const instance = await getMsal()
  const account = await getAccount()
  await instance.logoutRedirect(account ? { account } : undefined)
}

/** Token para Microsoft Graph; pide consentimiento solo si hace falta. */
export async function getToken(scopes: string[]): Promise<string> {
  const instance = await getMsal()
  const account = await getAccount()
  if (!account) {
    // Sin sesión: redirige a Microsoft (la página regresa ya conectada).
    await signIn()
    throw new Error('Redirigiendo a Microsoft para iniciar sesión…')
  }
  try {
    const silent = await instance.acquireTokenSilent({ scopes, account })
    return silent.accessToken
  } catch (error) {
    // CUALQUIER fallo en silencio (falta de consentimiento, timed_out del
    // iframe oculto, red lenta…) cae al flujo de página completa, que es
    // el que funciona de forma confiable en la Chromebook. Antes solo se
    // caía con InteractionRequiredAuthError y errores como timed_out
    // dejaban el botón muerto.
    const msg = error instanceof Error ? error.message : String(error)
    if (msg.includes('interaction_in_progress')) throw friendlyAuthError(error)
    try {
      await instance.acquireTokenRedirect({ scopes })
    } catch (e) {
      throw friendlyAuthError(e)
    }
    throw new Error('Redirigiendo a Microsoft para autorizar…')
  }
}
