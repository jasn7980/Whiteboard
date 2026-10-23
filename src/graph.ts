import { TIME_ZONE } from './config'

const GRAPH = 'https://graph.microsoft.com/v1.0'

async function graph<T>(
  token: string,
  path: string,
  init?: RequestInit & { body?: string },
): Promise<T> {
  const response = await fetch(`${GRAPH}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(init?.headers ?? {}),
    },
  })
  if (!response.ok) {
    const text = await response.text()
    throw new Error(`Graph ${response.status}: ${text.slice(0, 300)}`)
  }
  if (response.status === 204) return undefined as T
  return (await response.json()) as T
}

// --- Microsoft To Do --------------------------------------------------------

type TodoListsResponse = {
  value: { id: string; displayName: string; wellknownListName?: string }[]
}

let cachedListId: string | null = null

export async function getDefaultTaskListId(token: string): Promise<string> {
  if (cachedListId) return cachedListId
  const lists = await graph<TodoListsResponse>(token, '/me/todo/lists')
  const preferred =
    lists.value.find((l) => l.wellknownListName === 'defaultList') ?? lists.value[0]
  if (!preferred) throw new Error('No se encontró ninguna lista de tareas')
  cachedListId = preferred.id
  return cachedListId
}

export async function createTodoTask(
  token: string,
  title: string,
  body?: string,
): Promise<{ listId: string; taskId: string }> {
  const listId = await getDefaultTaskListId(token)
  const task = await graph<{ id: string }>(
    token,
    `/me/todo/lists/${encodeURIComponent(listId)}/tasks`,
    {
      method: 'POST',
      body: JSON.stringify({
        title,
        ...(body
          ? { body: { content: body, contentType: 'text' } }
          : {}),
      }),
    },
  )
  return { listId, taskId: task.id }
}

export async function setTodoTaskDone(
  token: string,
  listId: string,
  taskId: string,
  done: boolean,
): Promise<void> {
  await graph(
    token,
    `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
    {
      method: 'PATCH',
      body: JSON.stringify({ status: done ? 'completed' : 'notStarted' }),
    },
  )
}

// --- Outlook Calendar -------------------------------------------------------

export async function createCalendarEvent(
  token: string,
  subject: string,
  startLocalIso: string,
  endLocalIso: string,
): Promise<string> {
  const event = await graph<{ id: string }>(token, '/me/events', {
    method: 'POST',
    body: JSON.stringify({
      subject,
      start: { dateTime: startLocalIso, timeZone: TIME_ZONE },
      end: { dateTime: endLocalIso, timeZone: TIME_ZONE },
    }),
  })
  return event.id
}

// --- OneNote ----------------------------------------------------------------

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

/**
 * Crea una página en la sección por defecto del cuaderno por defecto
 * del usuario (POST /me/onenote/pages). Devuelve el id de la página.
 */
export async function createOneNotePage(
  token: string,
  title: string,
  body: string,
): Promise<string> {
  const html = [
    '<!DOCTYPE html>',
    '<html>',
    `<head><title>${escapeHtml(title)}</title></head>`,
    '<body>',
    body ? `<p>${escapeHtml(body)}</p>` : '',
    '<p><i>Creado desde el Tablero Whiteboard</i></p>',
    '</body>',
    '</html>',
  ].join('\n')
  const page = await graph<{ id: string }>(token, '/me/onenote/pages', {
    method: 'POST',
    headers: { 'Content-Type': 'text/html' },
    body: html,
  })
  return page.id
}
