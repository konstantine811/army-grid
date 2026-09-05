import {
  CacheKeys,
  invalidateDataCache,
  invalidatePersonnelCaches,
} from './data/idbDataCache'
import {
  clearAuthToken,
  emitAuthLogout,
  getAuthToken,
  type AuthSession,
  type AuthUser,
  type RegisteredUser,
} from './auth/authTypes'
import { showAppToast, showBackendBlockedToast } from './shared/appToast'
import { dataUrlToUint8Array } from './shared/browserExport'
import { createPhotoThumbnailDataUrl } from './pages/personnel/photoCompression'
import { apiRequestPool, type ApiRequestPriority } from './apiRequestPool'
import { measuredFetch } from './performance/performanceMonitor'

const resolveApiBaseUrl = () => {
  const fromEnv = import.meta.env.VITE_API_BASE_URL
  if (typeof fromEnv === 'string' && fromEnv.trim()) return fromEnv.trim()

  // Through Vite proxy (/api → :4000): same origin, works on LAN / HTTPS
  // without CORS or mixed-content blocks.
  if (typeof window !== 'undefined') {
    const host = window.location.hostname
    if (
      window.location.protocol === 'https:' ||
      (host && host !== 'localhost' && host !== '127.0.0.1')
    ) {
      return '/api'
    }
  }

  return 'http://127.0.0.1:4000'
}

/** Resolve per call — hostname differs on localhost vs phone/LAN. */
const apiBaseUrl = () => resolveApiBaseUrl()

type JsonRecord = Record<string, unknown>

export type BackendHealth = {
  ok: boolean
  database: 'connected' | 'unavailable'
  error?: string
}

export type BackendEjournalImportSheet = {
  id: string
  batchId: string
  name: string
  sheetIndex: number
  columnCount: number
  rowCount: number
  columns?: unknown
  createdAt: string
  updatedAt: string
}

export type BackendEjournalImport = {
  id: string
  kind: string
  name: string
  sourceFileName?: string | null
  sheetCount: number
  totalRows: number
  status: string
  notes?: string | null
  sheets: BackendEjournalImportSheet[]
  createdAt: string
  updatedAt: string
}

export type BackendEjournalImportRow = {
  id: string
  sheetId: string
  excelRowNumber?: number | null
  values: JsonRecord
  createdAt: string
}

export type BackendEjournalLiveVersion = {
  id: string
  unitLabel: string
  version: number
  sourceFileName?: string | null
  asOfDate?: string | null
  sha256: string
  byteSize: number
  baseVersionId?: string | null
  changeProtocol?: JsonRecord | null
  sourcePbFileName?: string | null
  sourcePbSha256?: string | null
  notes?: string | null
  createdByEmail?: string | null
  createdAt: string
  fileBase64?: string
}

export type BackendEjournalLiveState = {
  unitLabel: string
  current: BackendEjournalLiveVersion | null
  versions: BackendEjournalLiveVersion[]
}

export type BackendEjournalPbSource = {
  id: string
  unitLabel: string
  sourceFileName?: string | null
  asOfDate?: string | null
  sha256: string
  byteSize: number
  notes?: string | null
  createdByEmail?: string | null
  createdAt: string
  fileBase64?: string
}

export type BackendEjournalPbState = {
  unitLabel: string
  current: BackendEjournalPbSource | null
  items: BackendEjournalPbSource[]
}

export type BackendEjournalManualOperation = {
  id: string
  unitLabel: string
  status: 'draft' | 'applied' | 'cancelled'
  decision: 'pending' | 'accepted'
  input: JsonRecord
  baseVersionId: string
  appliedVersionId?: string | null
  appliedAt?: string | null
  createdByUserId?: string | null
  createdByEmail?: string | null
  createdByDisplayName?: string | null
  createdAt: string
  updatedAt: string
}

export type BackendEjournalSheetRows = {
  total: number
  limit: number
  offset: number
  columns: unknown
  items: BackendEjournalImportRow[]
}

export type BackendPersonnelRosterLatest = {
  importId: string
  importName: string
  sourceFileName?: string | null
  createdAt: string
  sheet: BackendEjournalImportSheet | null
  rows: BackendEjournalImportRow[]
}

export type BackendPersonnelRosterVersion = {
  importId: string
  createdAt: string
  sheetUpdatedAt: string | null
  rowCount: number
}

export type EjournalRowActionType =
  | 'MEDICAL'
  | 'LEAVE'
  | 'BUSINESS_TRIP'
  | 'AWOL'
  | 'CAPTIVITY'
  | 'EXCLUSION'
  | 'RETURNED'
  | 'POSITION_CHANGE'
  | 'RANK_CHANGE'

export type BackendPersonPhoto = {
  id: string
  personExternalId: string
  fileName?: string | null
  mimeType?: string | null
  photoData: string
  hasFile?: boolean
  hasThumbnail?: boolean
  crop?: JsonRecord | null
  createdAt: string
  updatedAt: string
}

const photoBlobToDataUrl = (blob: Blob) =>
  new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result || ''))
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(blob)
  })

const fetchPersonPhotoData = async (
  personExternalId: string,
  thumbnail = false,
) => {
  const query = thumbnail ? '?thumbnail=1' : ''
  return apiRequestPool.run(async () => {
    const response = await measuredFetch(
      `${apiBaseUrl()}/ejournals/personnel/photos/${encodeURIComponent(personExternalId)}/file${query}`,
      { headers: authHeaders() },
    )
    if (response.status === 404) return ''
    if (!response.ok) throw new Error(await parseApiError(response))
    return photoBlobToDataUrl(await response.blob())
  })
}

export type BackendPersonQuestionnaire = {
  id: string
  personExternalId: string
  fileName?: string | null
  mimeType?: string | null
  fileData: string
  createdAt: string
  updatedAt: string
}

export type BackendPersonQuestionnaireMeta = {
  personExternalId: string
  fileName?: string | null
}

export type BackendPersonDocument = {
  id: string
  personExternalId: string
  personName?: string | null
  type: string
  title: string
  status?: string | null
  fields?: JsonRecord | null
  workflow?: JsonRecord | null
  files?: JsonRecord | null
  createdByUserId?: string | null
  createdByEmail?: string | null
  createdByDisplayName?: string | null
  createdAt: string
  updatedAt: string
}

export type DocumentSignatoryBlockType = 'SIGNER' | 'APPROVAL'

export type BackendAnketaCellEdit = {
  id: string
  sheetId: string
  gid: string
  rowNumber: number
  columnId: string
  value: string
  externalId?: string | null
  fullName?: string | null
  updatedAt: string
  createdAt: string
}

export type BackendAnketaSheetSnapshot = {
  id: string
  sheetId: string
  gid: string
  payload: JsonRecord
  source: string
  sourceLabel?: string | null
  fetchedAt: string
  updatedAt: string
  createdAt: string
}

export type WorkTaskStatus = "open" | "done" | "irrelevant"
export type WorkTaskCategory = "ubd_status" | "ubd_send" | "tvk" | "other"

export type BackendWorkTaskComment = {
  id: string
  taskId: string
  body: string
  createdAt: string
}

export type BackendWorkTask = {
  id: string
  ownerUserId: string
  ownerEmail: string
  title: string
  personName?: string | null
  category: WorkTaskCategory | string
  location?: string | null
  status: WorkTaskStatus | string
  comments: BackendWorkTaskComment[]
  createdAt: string
  updatedAt: string
}

export type BackendDocumentSignatoryPreset = {
  id: string
  label: string
  blockType: DocumentSignatoryBlockType
  title: string
  rank: string
  fullName: string
  signatureData?: string | null
  signatureFileName?: string | null
  signatureMimeType?: string | null
  showDate: boolean
  documentTypes: string[]
  sortOrder: number
  createdAt: string
  updatedAt: string
}

export type SaveDocumentSignatoryPreset = Omit<
  BackendDocumentSignatoryPreset,
  'id' | 'createdAt' | 'updatedAt'
>

export type BackendPersonnelProfile = {
  externalId: string | null
  fullName: string | null
  person: JsonRecord | null
  ejournal: {
    importId: string | null
    oosRow: JsonRecord | null
    historicalOosRow?: JsonRecord | null
    absentRows: JsonRecord[]
  }
  roster: {
    importId: string
    row: JsonRecord
  } | null
  photo: JsonRecord | null
  questionnaire: JsonRecord | null
  documents: BackendPersonDocument[]
  changeLogs: JsonRecord[]
  exitPeriods: {
    absences: JsonRecord[]
    openAbsences: JsonRecord[]
    closedAbsences: JsonRecord[]
    locationPeriods: JsonRecord[]
    servicePeriods: JsonRecord[]
    rosterEvents: JsonRecord[]
    temporaryArrivals: JsonRecord[]
    fighterStatus: JsonRecord | null
    absentSheetRows: JsonRecord[]
    oosExitFields: JsonRecord
    hasAny: boolean
  }
}

export type AiQuestionnaireOcrField = {
  key: string
  label: string
  value: string
  confidence: 'high' | 'medium' | 'low' | string
}

export type AiQuestionnaireOcrResult = {
  model: string
  fileName?: string | null
  pageNumber?: string | null
  fields: AiQuestionnaireOcrField[]
}

export type DiskQuestionnaireMatchLevel = 'fio' | 'fi' | 'surname' | 'callsign'

export type DiskQuestionnaireMatch = {
  relativePath: string
  fileName: string
  matchLevel: DiskQuestionnaireMatchLevel
  score: number
  callSign: string
}

export type DiskQuestionnaireSearchPersonResult = {
  externalId: string
  fullName: string
  callSign: string
  missingQuestionnaire: boolean
  missingPhoto: boolean
  matches: DiskQuestionnaireMatch[]
}

export type DiskQuestionnaireSearchResult = {
  root: string
  scannedFiles: number
  matchedPeople: number
  people: DiskQuestionnaireSearchPersonResult[]
}

export type OverviewStatus =
  | 'ON_DUTY'
  | 'BUSINESS_TRIP'
  | 'LEAVE'
  | 'MEDICAL'
  | 'AWOL'
  | 'CAPTIVITY'
  | 'MISSING'
  | 'OTHER'

export type BackendPersonnelOverviewRow = {
  id: string
  externalId: string
  name: string
  rank: string
  unit: string
  positionTitle?: string
  status: OverviewStatus | string
  statusLabel: string
  fighterDirection?: string
  fighterEntryDate?: string
  fighterExitDate?: string
  fighterReturnDate?: string
  fighterTotalDays?: string
  fighterStatus?: string
  validFrom: string | null
  days: number | null
  plannedReturn: string | null
  place: string
  updatedAt: string
  /** Є в останній Штатці / ранковому «Загальному списку». */
  inStaff?: boolean
  /** Батальйон «нова» в останній Штатці — той самий набір, що БЧС. */
  inNovaStaff?: boolean
  /** Колонка A Штатки: нова / стара / інші пункти. */
  battalion?: string
  /** Є в останньому імпорті ЕЖООС. */
  fromEjoos?: boolean
  /** Статус з колонки «Статус» у Штатці (не ЕЖООС). */
  staffStatus?: OverviewStatus | string
  staffStatusLabel?: string
  /** Значення всіх колонок Штатки для перемикання в «Колонки». */
  staffSheetColumns?: Record<string, string>
}

export type BackendPersonnelOverview = {
  importId: string | null
  importName?: string
  metrics: {
    total: number
    onDuty: number
    businessTrip: number
    leave: number
    medical: number
    awol: number
    other: number
  }
  units: string[]
  rows: BackendPersonnelOverviewRow[]
  critical: Array<{
    id: string
    severity: 'danger' | 'warning' | 'info' | string
    text: string
    status: string
    days: number | null
    daysToReturn: number | null
  }>
  todayChanges: {
    total: number
    onDuty: number
    businessTrip: number
    leave: number
    medical: number
    awol: number
    other: number
  }
  todayUpdates: number
  cache?: {
    generatedAt: string
    expiresAt: string
    hit: boolean
  }
}

const authHeaders = (): HeadersInit => {
  const token = getAuthToken()
  return token ? { Authorization: `Bearer ${token}` } : {}
}

const parseApiError = async (response: Response) => {
  const text = await response.text()
  if (response.status === 401) {
    clearAuthToken()
    emitAuthLogout()
  }
  if (response.status === 413) {
    return 'Файл або запит завеликий для API. Спробуйте менший PDF або перезапустіть backend з більшим REQUEST_BODY_LIMIT.'
  }
  try {
    const json = JSON.parse(text) as { message?: string | string[] }
    if (Array.isArray(json.message)) return json.message.join(', ')
    if (typeof json.message === 'string' && json.message.trim()) return json.message
  } catch {
    // plain text
  }
  return text || `API request failed: ${response.status}`
}

type ApiRequestInit = RequestInit & {
  suppressErrorToast?: boolean
  poolPriority?: ApiRequestPriority
}

const requestPriority = (method?: string): ApiRequestPriority =>
  ['POST', 'PUT', 'PATCH', 'DELETE'].includes((method || 'GET').toUpperCase())
    ? 'high'
    : 'normal'

async function request<T>(path: string, options: ApiRequestInit = {}): Promise<T> {
  const { suppressErrorToast, poolPriority, ...fetchOptions } = options
  return apiRequestPool.run(
    async () => {
      const response = await measuredFetch(`${apiBaseUrl()}${path}`, {
        ...fetchOptions,
        headers: {
          'Content-Type': 'application/json',
          ...authHeaders(),
          ...fetchOptions.headers,
        },
      })

      if (!response.ok) {
        const message = await parseApiError(response)
        const method = (fetchOptions.method || 'GET').toUpperCase()
        const isWrite = ['POST', 'PUT', 'PATCH', 'DELETE'].includes(method)
        if (response.status === 403) {
          showBackendBlockedToast(message)
        } else if (
          isWrite &&
          response.status >= 400 &&
          response.status !== 401 &&
          !suppressErrorToast
        ) {
          showAppToast({
            title: 'Помилка запису',
            description: message,
            variant: response.status >= 500 ? 'CRITICAL' : 'WARNING',
          })
        }
        throw new Error(message)
      }

      if (response.status === 204) return undefined as T
      return response.json() as Promise<T>
    },
    {
      priority: poolPriority ?? requestPriority(fetchOptions.method),
      signal: fetchOptions.signal ?? undefined,
    },
  )
}

/** GET that returns null on 404 instead of throwing (no error toast). */
async function requestGetOptional<T>(path: string): Promise<T | null> {
  return apiRequestPool.run(async () => {
    const response = await measuredFetch(`${apiBaseUrl()}${path}`, {
      headers: {
        'Content-Type': 'application/json',
        ...authHeaders(),
      },
    })

    if (response.status === 404) return null
    if (!response.ok) {
      const message = await parseApiError(response)
      throw new Error(message)
    }
    if (response.status === 204) return null
    return response.json() as Promise<T>
  })
}

async function requestBinary<T>(path: string, body: Blob, options: RequestInit = {}): Promise<T> {
  return apiRequestPool.run(
    async () => {
      const response = await measuredFetch(`${apiBaseUrl()}${path}`, {
        ...options,
        body,
        headers: {
          ...authHeaders(),
          ...options.headers,
        },
      })

      if (!response.ok) {
        const text = await response.text()
        if (response.status === 401) {
          clearAuthToken()
          emitAuthLogout()
        }
        let message = text || `API request failed: ${response.status}`
        if (response.status === 413) {
          message =
            'PDF завеликий для API. Спробуйте стиснути файл або збільшити QUESTIONNAIRE_BODY_LIMIT на backend.'
        } else {
          try {
            const json = JSON.parse(text) as { message?: string | string[] }
            if (Array.isArray(json.message)) message = json.message.join(', ')
            else if (typeof json.message === 'string' && json.message.trim())
              message = json.message
          } catch {
            /* plain text */
          }
        }
        if (response.status === 403) {
          showBackendBlockedToast(message)
        } else if (response.status !== 401) {
          showAppToast({
            title: 'Помилка запису',
            description: message,
            variant: response.status >= 500 ? 'CRITICAL' : 'WARNING',
          })
        }
        throw new Error(message)
      }

      return response.json() as Promise<T>
    },
    {
      priority: 'high',
      signal: options.signal ?? undefined,
    },
  )
}

export const api = {
  get baseUrl() {
    return apiBaseUrl()
  },

  register(body: { email: string; password: string; displayName: string }) {
    return request<AuthSession>('/auth/register', {
      method: 'POST',
      body: JSON.stringify(body),
    })
  },

  login(body: { email: string; password: string }) {
    return request<AuthSession>('/auth/login', {
      method: 'POST',
      body: JSON.stringify(body),
    })
  },

  getAuthMe() {
    return request<AuthUser>('/auth/me')
  },

  updateOwnProfile(payload: {
    displayName?: string
    nickname?: string
    photoData?: string
    linkedPersonExternalId?: string | null
    linkedPersonFullName?: string | null
  }) {
    return request<AuthUser>('/auth/me/profile', {
      method: 'PATCH',
      body: JSON.stringify(payload),
    })
  },

  refreshAuth() {
    return request<AuthSession>('/auth/refresh', { method: 'POST' })
  },

  listAuthUsers() {
    return request<RegisteredUser[]>('/auth/users')
  },

  setUserAccess(userId: string, accessGranted: boolean) {
    return request<RegisteredUser>(`/auth/users/${encodeURIComponent(userId)}/access`, {
      method: 'PATCH',
      body: JSON.stringify({ accessGranted }),
    })
  },

  setUserPermissions(userId: string, writePermissions: string[]) {
    return request<RegisteredUser>(
      `/auth/users/${encodeURIComponent(userId)}/permissions`,
      {
        method: 'PATCH',
        body: JSON.stringify({ writePermissions }),
      },
    )
  },

  getHealth() {
    return request<BackendHealth>('/health')
  },

  listEjournalImports(options: { signal?: AbortSignal } = {}) {
    return request<BackendEjournalImport[]>('/ejournals/imports', {
      signal: options.signal,
    })
  },

  listEjournalSheetRows(
    sheetId: string,
    options: { limit?: number; offset?: number; signal?: AbortSignal } = {},
  ) {
    const searchParams = new URLSearchParams({
      limit: String(options.limit ?? 500),
      offset: String(options.offset ?? 0),
    })

    return request<BackendEjournalSheetRows>(
      `/ejournals/sheets/${sheetId}/rows?${searchParams.toString()}`,
      { signal: options.signal },
    )
  },

  listBchsImports() {
    return request<BackendEjournalImport[]>('/bchs/imports')
  },

  listBchsSheetRows(sheetId: string, options: { limit?: number; offset?: number } = {}) {
    const searchParams = new URLSearchParams({
      limit: String(options.limit ?? 500),
      offset: String(options.offset ?? 0),
    })

    return request<BackendEjournalSheetRows>(`/bchs/sheets/${sheetId}/rows?${searchParams.toString()}`)
  },

  deleteBchsImport(batchId: string) {
    return request<{ id: string; deleted: true }>(`/bchs/imports/${batchId}`, {
      method: 'DELETE',
    })
  },

  importBchsWorkbook(payload: {
    name: string
    sourceFileName?: string
    notes?: string
    sheets: Array<{
      name: string
      sheetIndex: number
      columns: Array<{ key: string; label: string; order: number; originalIndex?: number; letter?: string }>
      rows: Array<{ excelRowNumber?: number; values: JsonRecord }>
    }>
  }) {
    return request<BackendEjournalImport>('/bchs/import-workbook', {
      method: 'POST',
      body: JSON.stringify(payload),
    })
  },

  updateEjournalRowValues(
    rowId: string,
    values: JsonRecord,
    options?: { suppressErrorToast?: boolean },
  ) {
    return request<BackendEjournalImportRow>(`/ejournals/rows/${rowId}`, {
      method: 'PATCH',
      body: JSON.stringify({ values, actor: 'operator' }),
      suppressErrorToast: options?.suppressErrorToast,
    }).then(async (result) => {
      await invalidateDataCache(
        'ejournal:sheet-rows:',
        CacheKeys.rosterLatest,
        CacheKeys.overview,
        CacheKeys.personnelDataset,
      )
      return result
    })
  },

  aiQuestionnaireOcr(payload: {
    imageData: string
    fileName?: string
    pageNumber?: string
  }) {
    return request<AiQuestionnaireOcrResult>('/ejournals/questionnaire-ai-ocr', {
      method: 'POST',
      body: JSON.stringify(payload),
    })
  },

  async getPersonPhoto(personExternalId: string) {
    const photo = await request<BackendPersonPhoto | null>(
      `/ejournals/personnel/photos/${encodeURIComponent(personExternalId)}`,
    )
    if (!photo || photo.photoData || !photo.hasFile) return photo
    return {
      ...photo,
      photoData: await fetchPersonPhotoData(personExternalId),
    }
  },

  getPersonPhotoThumbnail(personExternalId: string) {
    return fetchPersonPhotoData(personExternalId, true)
  },

  getPersonnelOverview(
    options: {
      limit?: number
      offset?: number
      force?: boolean
      signal?: AbortSignal
    } = {},
  ) {
    const searchParams = new URLSearchParams()
    if (options.limit != null) searchParams.set('limit', String(options.limit))
    if (options.offset != null) searchParams.set('offset', String(options.offset))
    if (options.force) searchParams.set('force', '1')
    const query = searchParams.toString()
    return request<BackendPersonnelOverview>(
      `/ejournals/personnel/overview${query ? `?${query}` : ''}`,
      { signal: options.signal },
    )
  },

  listPersonPhotos() {
    return request<Array<{
      personExternalId: string
      photoData: string
      hasFile?: boolean
      hasThumbnail?: boolean
    }>>(
      '/ejournals/personnel/photos',
    )
  },

  async upsertPersonPhoto(personExternalId: string, payload: {
    photoData: string
    thumbnailData?: string
    fileName?: string
    mimeType?: string
    crop?: JsonRecord
  }) {
    const thumbnailData =
      payload.thumbnailData ||
      (await createPhotoThumbnailDataUrl(payload.photoData).catch(() => undefined))
    return request<BackendPersonPhoto>(`/ejournals/personnel/photos/${encodeURIComponent(personExternalId)}`, {
      method: 'PATCH',
      body: JSON.stringify({ ...payload, thumbnailData }),
    })
  },

  deletePersonPhoto(personExternalId: string) {
    return request<{ personExternalId: string; deleted: boolean }>(
      `/ejournals/personnel/photos/${encodeURIComponent(personExternalId)}`,
      { method: 'DELETE' },
    )
  },

  getPersonQuestionnaire(personExternalId: string) {
    return request<BackendPersonQuestionnaire | null>(
      `/ejournals/personnel/questionnaires/${encodeURIComponent(personExternalId)}`,
    )
  },

  getPersonQuestionnaireFileUrl(
    personExternalId: string,
    fileName?: string,
    download = false,
  ) {
    const params = new URLSearchParams()
    if (fileName) params.set('fileName', fileName)
    if (download) params.set('download', '1')
    const query = params.toString()
    return `${apiBaseUrl()}/ejournals/personnel/questionnaires/${encodeURIComponent(personExternalId)}/file${
      query ? `?${query}` : ''
    }`
  },

  async createPersonQuestionnairePreviewUrl(
    personExternalId: string,
    fileName: string,
  ) {
    try {
      return await this.getPersonQuestionnaireObjectUrl(
        personExternalId,
        fileName,
      )
    } catch {
      const questionnaire = await this.getPersonQuestionnaire(personExternalId)
      if (!questionnaire?.fileData) {
        throw new Error('Немає PDF анкети для перегляду.')
      }
      const bytes = dataUrlToUint8Array(questionnaire.fileData)
      return URL.createObjectURL(
        new Blob([bytes], {
          type: questionnaire.mimeType || 'application/pdf',
        }),
      )
    }
  },

  async fetchPersonQuestionnaireFile(
    personExternalId: string,
    fileName?: string,
    download = false,
  ) {
    return apiRequestPool.run(
      async () => {
        const response = await measuredFetch(
          this.getPersonQuestionnaireFileUrl(personExternalId, fileName, download),
          { headers: { ...authHeaders() } },
        )
        if (!response.ok) {
          throw new Error(await parseApiError(response))
        }
        return response.blob()
      },
      { priority: 'high' },
    )
  },

  async getPersonQuestionnaireObjectUrl(
    personExternalId: string,
    fileName?: string,
  ) {
    const blob = await this.fetchPersonQuestionnaireFile(personExternalId, fileName)
    return URL.createObjectURL(blob)
  },

  listPersonQuestionnaires(options: { signal?: AbortSignal } = {}) {
    return request<BackendPersonQuestionnaireMeta[]>(
      '/ejournals/personnel/questionnaires',
      { signal: options.signal },
    )
  },

  searchQuestionnairesOnDisk(payload: {
    people: Array<{
      externalId: string
      fullName: string
      callSign?: string
      missingQuestionnaire?: boolean
      missingPhoto?: boolean
    }>
    refreshIndex?: boolean
  }) {
    return request<DiskQuestionnaireSearchResult>(
      '/ejournals/personnel/questionnaire-disk/search',
      {
        method: 'POST',
        body: JSON.stringify(payload),
        poolPriority: 'normal',
      },
    )
  },

  async getDiskQuestionnaireFile(relativePath: string) {
    return apiRequestPool.run(
      async () => {
        const response = await measuredFetch(
          `${apiBaseUrl()}/ejournals/personnel/questionnaire-disk/file`,
          {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              ...authHeaders(),
            },
            body: JSON.stringify({ relativePath }),
          },
        )
        if (!response.ok) {
          if (response.status === 401) {
            clearAuthToken()
            emitAuthLogout()
          }
          const text = await response.text()
          let message = text || `API request failed: ${response.status}`
          try {
            const parsed = JSON.parse(text) as { message?: string }
            if (parsed?.message) message = parsed.message
          } catch {
            // keep raw text
          }
          throw new Error(message)
        }
        return response.blob()
      },
      { priority: 'high' },
    )
  },

  revealDiskQuestionnaireInFinder(relativePath: string) {
    return request<{ revealed: boolean; fileName: string }>(
      '/ejournals/personnel/questionnaire-disk/reveal',
      {
        method: 'POST',
        body: JSON.stringify({ relativePath }),
      },
    )
  },

  confirmDiskQuestionnaire(
    personExternalId: string,
    relativePath: string,
    fileName?: string,
    options?: {
      suppressErrorToast?: boolean
      poolPriority?: ApiRequestPriority
    },
  ) {
    return request<BackendPersonQuestionnaire>(
      `/ejournals/personnel/questionnaire-disk/${encodeURIComponent(personExternalId)}/confirm`,
      {
        method: 'POST',
        body: JSON.stringify({
          relativePath,
          ...(fileName ? { fileName } : {}),
        }),
        suppressErrorToast: options?.suppressErrorToast,
        poolPriority: options?.poolPriority,
      },
    )
  },

  upsertPersonQuestionnaire(personExternalId: string, payload: {
    fileData: string
    fileName?: string
    mimeType?: string
  }) {
    return request<BackendPersonQuestionnaire>(
      `/ejournals/personnel/questionnaires/${encodeURIComponent(personExternalId)}`,
      {
        method: 'PATCH',
        body: JSON.stringify(payload),
      },
    )
  },

  upsertPersonQuestionnaireFile(personExternalId: string, file: File) {
    return requestBinary<BackendPersonQuestionnaire>(
      `/ejournals/personnel/questionnaires/${encodeURIComponent(personExternalId)}/file`,
      file,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/pdf',
          'X-File-Name': encodeURIComponent(file.name),
        },
      },
    )
  },

  deletePersonQuestionnaire(personExternalId: string) {
    return request<{ personExternalId: string; deleted: boolean }>(
      `/ejournals/personnel/questionnaires/${encodeURIComponent(personExternalId)}`,
      { method: 'DELETE' },
    )
  },

  listPersonDocuments(
    personExternalId: string,
    options: { full?: boolean } = {},
  ) {
    const query = options.full ? '?full=1' : ''
    return request<BackendPersonDocument[]>(
      `/ejournals/personnel/${encodeURIComponent(personExternalId)}/documents${query}`,
    )
  },

  getPersonnelProfile(personExternalId: string, fullName?: string) {
    const params = new URLSearchParams()
    if (fullName?.trim()) params.set('fullName', fullName.trim())
    const query = params.toString()
    return requestGetOptional<BackendPersonnelProfile>(
      `/ejournals/personnel/${encodeURIComponent(personExternalId)}/profile${query ? `?${query}` : ''}`,
    )
  },

  listAllPersonDocuments(options: { signal?: AbortSignal } = {}) {
    return request<BackendPersonDocument[]>('/ejournals/personnel/documents', {
      signal: options.signal,
    })
  },

  createPersonDocument(
    personExternalId: string,
    payload: {
      type: string
      title: string
      status?: string
      fields?: JsonRecord
      workflow?: JsonRecord
      files?: JsonRecord
    },
    options?: { suppressErrorToast?: boolean },
  ) {
    return request<BackendPersonDocument>(
      `/ejournals/personnel/${encodeURIComponent(personExternalId)}/documents`,
      {
        method: 'POST',
        body: JSON.stringify(payload),
        suppressErrorToast: options?.suppressErrorToast,
      },
    ).then(async (result) => {
      await invalidateDataCache(CacheKeys.documentsAll, CacheKeys.overview)
      return result
    })
  },

  updatePersonDocument(
    personExternalId: string,
    documentId: string,
    payload: {
      title?: string
      status?: string
      fields?: JsonRecord
      workflow?: JsonRecord
      files?: JsonRecord
    },
    options?: { suppressErrorToast?: boolean },
  ) {
    return request<BackendPersonDocument>(
      `/ejournals/personnel/${encodeURIComponent(personExternalId)}/documents/${encodeURIComponent(documentId)}`,
      {
        method: 'PATCH',
        body: JSON.stringify(payload),
        suppressErrorToast: options?.suppressErrorToast,
      },
    ).then(async (result) => {
      await invalidateDataCache(CacheKeys.documentsAll, CacheKeys.overview)
      return result
    })
  },

  deletePersonDocument(personExternalId: string, documentId: string) {
    return request<{ id: string; personExternalId: string; deleted: boolean }>(
      `/ejournals/personnel/${encodeURIComponent(personExternalId)}/documents/${encodeURIComponent(documentId)}`,
      { method: 'DELETE' },
    ).then(async (result) => {
      await invalidateDataCache(CacheKeys.documentsAll, CacheKeys.overview)
      return result
    })
  },

  listDocumentSignatories(documentType?: string) {
    const query = documentType
      ? `?documentType=${encodeURIComponent(documentType)}`
      : ''
    return request<BackendDocumentSignatoryPreset[]>(
      `/document-signatories${query}`,
    )
  },

  createDocumentSignatory(payload: SaveDocumentSignatoryPreset) {
    return request<BackendDocumentSignatoryPreset>('/document-signatories', {
      method: 'POST',
      body: JSON.stringify(payload),
    })
  },

  updateDocumentSignatory(
    id: string,
    payload: SaveDocumentSignatoryPreset,
  ) {
    return request<BackendDocumentSignatoryPreset>(
      `/document-signatories/${encodeURIComponent(id)}`,
      {
        method: 'PATCH',
        body: JSON.stringify(payload),
      },
    )
  },

  deleteDocumentSignatory(id: string) {
    return request<{ id: string; deleted: boolean }>(
      `/document-signatories/${encodeURIComponent(id)}`,
      { method: 'DELETE' },
    )
  },

  createEjournalRowAction(rowId: string, payload: {
    actionType: EjournalRowActionType
    validFrom?: string
    validTo?: string
    reason?: string
    place?: string
    note?: string
    positionIndex?: string
    positionTitle?: string
    rank?: string
  }) {
    return request<{ person: JsonRecord; actionType: EjournalRowActionType; result: JsonRecord }>(`/ejournals/rows/${rowId}/actions`, {
      method: 'POST',
      body: JSON.stringify({ ...payload, actor: 'operator' }),
    })
  },

  importEjournalWorkbook(payload: {
    name: string
    sourceFileName?: string
    notes?: string
    sheets: Array<{
      name: string
      sheetIndex: number
      columns: Array<{ key: string; label: string; order: number; originalIndex?: number; letter?: string }>
      rows: Array<{ excelRowNumber?: number; values: JsonRecord }>
    }>
  }) {
    return request<BackendEjournalImport>('/ejournals/import-workbook', {
      method: 'POST',
      body: JSON.stringify(payload),
    }).then(async (result) => {
      await invalidatePersonnelCaches()
      return result
    })
  },

  importPersonnelRoster(payload: {
    name: string
    sourceFileName?: string
    notes?: string
    sheets: Array<{
      name: string
      sheetIndex: number
      columns: Array<{ key: string; label: string; order: number; originalIndex?: number; letter?: string }>
      rows: Array<{ excelRowNumber?: number; values: JsonRecord }>
    }>
  }) {
    return request<BackendEjournalImport>('/ejournals/personnel/roster/import', {
      method: 'POST',
      body: JSON.stringify(payload),
    }).then(async (result) => {
      await invalidateDataCache(
        CacheKeys.rosterLatest,
        CacheKeys.overview,
        CacheKeys.personnelDataset,
      )
      return result
    })
  },

  getLatestPersonnelRoster(options: { signal?: AbortSignal } = {}) {
    return request<BackendPersonnelRosterLatest | null>(
      '/ejournals/personnel/roster/latest',
      { signal: options.signal },
    )
  },

  getLatestPersonnelRosterVersion(options: { signal?: AbortSignal } = {}) {
    return request<BackendPersonnelRosterVersion | null>(
      '/ejournals/personnel/roster/latest/version',
      { signal: options.signal },
    )
  },

  getEjournalLive(unitLabel = '1ПБ') {
    return requestGetOptional<BackendEjournalLiveState>(
      `/ejournals/live?unitLabel=${encodeURIComponent(unitLabel)}`,
    ).then(
      (state) =>
        state ?? {
          unitLabel,
          current: null,
          versions: [],
        },
    )
  },

  getEjournalLiveFile(versionId?: string, unitLabel = '1ПБ') {
    const params = new URLSearchParams()
    params.set('unitLabel', unitLabel)
    if (versionId) params.set('versionId', versionId)
    return request<BackendEjournalLiveVersion>(`/ejournals/live/file?${params.toString()}`)
  },

  seedEjournalLive(payload: {
    fileBase64: string
    sourceFileName?: string
    asOfDate?: string
    unitLabel?: string
    notes?: string
  }) {
    return request<BackendEjournalLiveVersion>('/ejournals/live/seed', {
      method: 'POST',
      body: JSON.stringify(payload),
    })
  },

  applyEjournalLive(payload: {
    baseVersionId: string
    fileBase64: string
    sourceFileName?: string
    asOfDate?: string
    unitLabel?: string
    sourcePbFileName?: string
    sourcePbSha256?: string
    changeProtocol?: JsonRecord
    appliedManualOperationIds?: string[]
    notes?: string
  }) {
    return request<BackendEjournalLiveVersion>('/ejournals/live/apply', {
      method: 'POST',
      body: JSON.stringify(payload),
    })
  },

  listEjournalManualOperations(unitLabel = '1ПБ') {
    return request<BackendEjournalManualOperation[]>(
      `/ejournals/manual-operations?unitLabel=${encodeURIComponent(unitLabel)}`,
    )
  },

  createEjournalManualOperation(payload: {
    unitLabel?: string
    input: JsonRecord
    baseVersionId: string
    decision?: 'pending' | 'accepted'
  }) {
    return request<BackendEjournalManualOperation>(
      '/ejournals/manual-operations',
      {
        method: 'POST',
        body: JSON.stringify(payload),
      },
    )
  },

  updateEjournalManualOperation(
    id: string,
    payload: {
      input?: JsonRecord
      baseVersionId?: string
      decision?: 'pending' | 'accepted'
    },
  ) {
    return request<BackendEjournalManualOperation>(
      `/ejournals/manual-operations/${encodeURIComponent(id)}`,
      {
        method: 'PATCH',
        body: JSON.stringify(payload),
      },
    )
  },

  cancelEjournalManualOperation(id: string) {
    return request<BackendEjournalManualOperation>(
      `/ejournals/manual-operations/${encodeURIComponent(id)}/cancel`,
      { method: 'POST' },
    )
  },

  rollbackEjournalLive(payload: {
    targetVersionId: string
    unitLabel?: string
    fileBase64?: string
    notes?: string
  }) {
    return request<BackendEjournalLiveVersion>('/ejournals/live/rollback', {
      method: 'POST',
      body: JSON.stringify(payload),
    })
  },

  getEjournalPbSources(unitLabel = '1ПБ') {
    return requestGetOptional<BackendEjournalPbState>(
      `/ejournals/pb?unitLabel=${encodeURIComponent(unitLabel)}`,
    ).then(
      (state) =>
        state ?? {
          unitLabel,
          current: null,
          items: [],
        },
    )
  },

  getEjournalPbFile(id?: string, unitLabel = '1ПБ') {
    const params = new URLSearchParams()
    params.set('unitLabel', unitLabel)
    if (id) params.set('id', id)
    return request<BackendEjournalPbSource>(`/ejournals/pb/file?${params.toString()}`)
  },

  uploadEjournalPb(payload: {
    fileBase64: string
    sourceFileName?: string
    asOfDate?: string
    unitLabel?: string
    notes?: string
  }) {
    return request<BackendEjournalPbSource>('/ejournals/pb/upload', {
      method: 'POST',
      body: JSON.stringify(payload),
    })
  },

  getEjournalOperatorSettings(unitLabel = '1ПБ') {
    return request<{
      unitLabel: string
      settings: JsonRecord | null
      updatedAt: string | null
    }>(`/ejournals/operator-settings?unitLabel=${encodeURIComponent(unitLabel)}`)
  },

  putEjournalOperatorSettings(payload: {
    unitLabel?: string
    settings: JsonRecord
  }) {
    return request<{
      unitLabel: string
      settings: JsonRecord | null
      updatedAt: string | null
    }>('/ejournals/operator-settings', {
      method: 'PUT',
      body: JSON.stringify(payload),
    })
  },

  getEjournalNormalized(unitLabel = '1ПБ') {
    return request<{
      unitLabel: string
      versionId: string | null
      asOfDate: string | null
      syncedAt: string | null
      counts: {
        persons: number
        absences: number
        timesheet: number
        arrivals?: number
        irrevocableLosses?: number
      }
      persons: Array<{
        personId: string
        fullName: string
        rank: string
        positionIndex: string
        dayCode: string
        isVacant: boolean
      }>
    }>(`/ejournals/normalized?unitLabel=${encodeURIComponent(unitLabel)}`)
  },

  syncEjournalNormalized(payload: {
    unitLabel?: string
    versionId?: string
    asOfDate?: string | null
    persons: Array<Record<string, unknown>>
    absences: Array<Record<string, unknown>>
    timesheet: Array<Record<string, unknown>>
    arrivals?: Array<Record<string, unknown>>
    irrevocableLosses?: Array<Record<string, unknown>>
  }) {
    return request<{
      unitLabel: string
      counts: {
        persons: number
        absences: number
        timesheet: number
        arrivals?: number
        irrevocableLosses?: number
      }
      syncedAt: string | null
    }>('/ejournals/normalized/sync', {
      method: 'POST',
      body: JSON.stringify(payload),
    })
  },

  listAnketaEdits(sheetId?: string, gid?: string) {
    const params = new URLSearchParams()
    if (sheetId?.trim()) params.set('sheetId', sheetId.trim())
    if (gid?.trim()) params.set('gid', gid.trim())
    const query = params.toString()
    return request<BackendAnketaCellEdit[]>(
      `/anketa/edits${query ? `?${query}` : ''}`,
    )
  },

  upsertAnketaCellEdit(payload: {
    rowNumber: number
    columnId: string
    value: string
    externalId?: string
    fullName?: string
    sheetId?: string
    gid?: string
  }) {
    return request<BackendAnketaCellEdit>('/anketa/edits', {
      method: 'PATCH',
      body: JSON.stringify({ ...payload, actor: 'operator' }),
    })
  },

  bulkUpsertAnketaEdits(
    items: Array<{
      rowNumber: number
      columnId: string
      value: string
      externalId?: string
      fullName?: string
      sheetId?: string
      gid?: string
    }>,
  ) {
    return request<{ count: number; items: BackendAnketaCellEdit[] }>(
      '/anketa/edits/bulk',
      {
        method: 'POST',
        body: JSON.stringify({ items, actor: 'operator' }),
      },
    )
  },

  async getAnketaSnapshot(sheetId?: string, gid?: string) {
    const params = new URLSearchParams()
    if (sheetId?.trim()) params.set('sheetId', sheetId.trim())
    if (gid?.trim()) params.set('gid', gid.trim())
    const query = params.toString()
    try {
      return await request<BackendAnketaSheetSnapshot>(
        `/anketa/snapshot${query ? `?${query}` : ''}`,
      )
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      if (/не знайдено|not found|404/i.test(message)) return null
      throw error
    }
  },

  putAnketaSnapshot(payload: {
    payload: JsonRecord
    source?: string
    sourceLabel?: string
    fetchedAt?: string
    sheetId?: string
    gid?: string
  }) {
    return request<BackendAnketaSheetSnapshot>('/anketa/snapshot', {
      method: 'PUT',
      body: JSON.stringify({ ...payload, actor: 'operator' }),
    })
  },

  listWorkTasks(status?: string, q?: string) {
    const params = new URLSearchParams()
    if (status?.trim()) params.set('status', status.trim())
    if (q?.trim()) params.set('q', q.trim())
    const query = params.toString()
    return request<BackendWorkTask[]>(
      `/work-tasks${query ? `?${query}` : ''}`,
    )
  },

  createWorkTask(payload: {
    title: string
    personName?: string
    category?: string
    location?: string
    firstComment?: string
  }) {
    return request<BackendWorkTask>('/work-tasks', {
      method: 'POST',
      body: JSON.stringify(payload),
    })
  },

  updateWorkTask(
    id: string,
    payload: {
      title?: string
      personName?: string
      category?: string
      location?: string
      status?: string
    },
  ) {
    return request<BackendWorkTask>(`/work-tasks/${encodeURIComponent(id)}`, {
      method: 'PATCH',
      body: JSON.stringify(payload),
    })
  },

  addWorkTaskComment(id: string, body: string) {
    return request<BackendWorkTask>(
      `/work-tasks/${encodeURIComponent(id)}/comments`,
      {
        method: 'POST',
        body: JSON.stringify({ body }),
      },
    )
  },

  deleteWorkTask(id: string) {
    return request<{ ok: boolean }>(`/work-tasks/${encodeURIComponent(id)}`, {
      method: 'DELETE',
    })
  },
}
