const API_BASE_URL = import.meta.env.VITE_API_BASE_URL ?? 'http://127.0.0.1:4000'

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
  crop?: JsonRecord | null
  createdAt: string
  updatedAt: string
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
  type: string
  title: string
  status?: string | null
  fields?: JsonRecord | null
  workflow?: JsonRecord | null
  files?: JsonRecord | null
  createdAt: string
  updatedAt: string
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
  status: OverviewStatus | string
  statusLabel: string
  validFrom: string | null
  days: number | null
  plannedReturn: string | null
  place: string
  updatedAt: string
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
}

async function request<T>(path: string, options: RequestInit = {}): Promise<T> {
  const response = await fetch(`${API_BASE_URL}${path}`, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      ...options.headers,
    },
  })

  if (!response.ok) {
    const text = await response.text()
    if (response.status === 413) {
      throw new Error('Файл або запит завеликий для API. Спробуйте менший PDF або перезапустіть backend з більшим REQUEST_BODY_LIMIT.')
    }
    throw new Error(text || `API request failed: ${response.status}`)
  }

  return response.json() as Promise<T>
}

async function requestBinary<T>(path: string, body: Blob, options: RequestInit = {}): Promise<T> {
  const response = await fetch(`${API_BASE_URL}${path}`, {
    ...options,
    body,
  })

  if (!response.ok) {
    const text = await response.text()
    if (response.status === 413) {
      throw new Error('PDF завеликий для API. Спробуйте стиснути файл або збільшити QUESTIONNAIRE_BODY_LIMIT на backend.')
    }
    throw new Error(text || `API request failed: ${response.status}`)
  }

  return response.json() as Promise<T>
}

export const api = {
  baseUrl: API_BASE_URL,

  getHealth() {
    return request<BackendHealth>('/health')
  },

  listEjournalImports() {
    return request<BackendEjournalImport[]>('/ejournals/imports')
  },

  listEjournalSheetRows(sheetId: string, options: { limit?: number; offset?: number } = {}) {
    const searchParams = new URLSearchParams({
      limit: String(options.limit ?? 500),
      offset: String(options.offset ?? 0),
    })

    return request<BackendEjournalSheetRows>(`/ejournals/sheets/${sheetId}/rows?${searchParams.toString()}`)
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

  updateEjournalRowValues(rowId: string, values: JsonRecord) {
    return request<BackendEjournalImportRow>(`/ejournals/rows/${rowId}`, {
      method: 'PATCH',
      body: JSON.stringify({ values, actor: 'operator' }),
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

  getPersonPhoto(personExternalId: string) {
    return request<BackendPersonPhoto | null>(`/ejournals/personnel/photos/${encodeURIComponent(personExternalId)}`)
  },

  getPersonnelOverview() {
    return request<BackendPersonnelOverview>('/ejournals/personnel/overview')
  },

  listPersonPhotos() {
    return request<Array<{ personExternalId: string; photoData: string }>>(
      '/ejournals/personnel/photos',
    )
  },

  upsertPersonPhoto(personExternalId: string, payload: {
    photoData: string
    fileName?: string
    mimeType?: string
    crop?: JsonRecord
  }) {
    return request<BackendPersonPhoto>(`/ejournals/personnel/photos/${encodeURIComponent(personExternalId)}`, {
      method: 'PATCH',
      body: JSON.stringify(payload),
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

  listPersonQuestionnaires() {
    return request<BackendPersonQuestionnaireMeta[]>('/ejournals/personnel/questionnaires')
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
      },
    )
  },

  async getDiskQuestionnaireFile(relativePath: string) {
    const response = await fetch(
      `${API_BASE_URL}/ejournals/personnel/questionnaire-disk/file`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ relativePath }),
      },
    )
    if (!response.ok) {
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

  confirmDiskQuestionnaire(personExternalId: string, relativePath: string) {
    return request<BackendPersonQuestionnaire>(
      `/ejournals/personnel/questionnaire-disk/${encodeURIComponent(personExternalId)}/confirm`,
      {
        method: 'POST',
        body: JSON.stringify({ relativePath }),
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

  listPersonDocuments(personExternalId: string) {
    return request<BackendPersonDocument[]>(
      `/ejournals/personnel/${encodeURIComponent(personExternalId)}/documents`,
    )
  },

  listAllPersonDocuments() {
    return request<BackendPersonDocument[]>('/ejournals/personnel/documents')
  },

  createPersonDocument(personExternalId: string, payload: {
    type: string
    title: string
    status?: string
    fields?: JsonRecord
    workflow?: JsonRecord
    files?: JsonRecord
  }) {
    return request<BackendPersonDocument>(
      `/ejournals/personnel/${encodeURIComponent(personExternalId)}/documents`,
      {
        method: 'POST',
        body: JSON.stringify(payload),
      },
    )
  },

  updatePersonDocument(personExternalId: string, documentId: string, payload: {
    title?: string
    status?: string
    fields?: JsonRecord
    workflow?: JsonRecord
    files?: JsonRecord
  }) {
    return request<BackendPersonDocument>(
      `/ejournals/personnel/${encodeURIComponent(personExternalId)}/documents/${encodeURIComponent(documentId)}`,
      {
        method: 'PATCH',
        body: JSON.stringify(payload),
      },
    )
  },

  deletePersonDocument(personExternalId: string, documentId: string) {
    return request<{ id: string; personExternalId: string; deleted: boolean }>(
      `/ejournals/personnel/${encodeURIComponent(personExternalId)}/documents/${encodeURIComponent(documentId)}`,
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
    })
  },

  getLatestPersonnelRoster() {
    return request<BackendPersonnelRosterLatest | null>('/ejournals/personnel/roster/latest')
  },
}
