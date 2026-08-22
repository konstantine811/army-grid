import {
  CacheKeys,
  invalidateDataCache,
  invalidatePersonnelCaches,
} from './data/idbDataCache'

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
    }).then(async (result) => {
      await invalidateDataCache(
        'ejournal:sheet-rows:',
        CacheKeys.rosterLatest,
        CacheKeys.overview,
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

  getPersonQuestionnaireFileUrl(
    personExternalId: string,
    fileName?: string,
    download = false,
  ) {
    const params = new URLSearchParams()
    if (fileName) params.set('fileName', fileName)
    if (download) params.set('download', '1')
    const query = params.toString()
    return `${API_BASE_URL}/ejournals/personnel/questionnaires/${encodeURIComponent(personExternalId)}/file${
      query ? `?${query}` : ''
    }`
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

  confirmDiskQuestionnaire(
    personExternalId: string,
    relativePath: string,
    fileName?: string,
  ) {
    return request<BackendPersonQuestionnaire>(
      `/ejournals/personnel/questionnaire-disk/${encodeURIComponent(personExternalId)}/confirm`,
      {
        method: 'POST',
        body: JSON.stringify({
          relativePath,
          ...(fileName ? { fileName } : {}),
        }),
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

  getPersonnelProfile(personExternalId: string, fullName?: string) {
    const params = new URLSearchParams()
    if (fullName?.trim()) params.set('fullName', fullName.trim())
    const query = params.toString()
    return request<BackendPersonnelProfile>(
      `/ejournals/personnel/${encodeURIComponent(personExternalId)}/profile${query ? `?${query}` : ''}`,
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
    ).then(async (result) => {
      await invalidateDataCache(CacheKeys.documentsAll, CacheKeys.overview)
      return result
    })
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
        'ejournal:sheet-rows:',
      )
      return result
    })
  },

  getLatestPersonnelRoster() {
    return request<BackendPersonnelRosterLatest | null>('/ejournals/personnel/roster/latest')
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
}
