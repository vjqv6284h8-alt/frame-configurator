import { useEffect, useMemo, useRef, useState } from 'react'
import type { ChangeEvent } from 'react'
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
import * as XLSX from 'xlsx'
import './App.css'

type SegmentType = 'FIX' | 'SL'
type OpeningDirection = 'left' | 'right'
type FrameSeries = '65' | '80'
type LeafBottomMode = 'standard' | 'freeway'
type FrameSystemType = 'single-track' | 'double-track' | 'triple-track' | 'quadruple-track'

type Profile = {
  id: string
  article: string
  name: string
  defaultLength: number
  sectionImageDataUrl?: string
}

type HardwareItem = {
  id: string
  article: string
  name: string
}

type ReadyHardwareItem = {
  id: string
  article: string
  name: string
}

type SpecItem = {
  key: string
  article: string
  profile: string
  zone: string
  partQuantity: number
  deductionMm: number
  additionMm: number
  totalLengthMm: number
  withWasteLengthMm: number
  stockLengthMm: number
  barsCount: number
  sectionImageDataUrl?: string
  isHardware?: boolean
}

type ActiveTab =
  | 'config'
  | 'profiles'
  | 'hardware'
  | 'ready-hardware'
  | 'assemblies'
  | 'nodes'
  | 'settings'

type CustomNodeFormula =
  | 'framePerimeter'
  | 'frameTop'
  | 'frameSidesBottom'
  | 'frameHeight'
  | 'panelWidth'
  | 'slPerimeterSingle'

type CustomNodeQuantityMode = 'fixed' | 'slidingCount' | 'segmentCount' | 'verticalJoints'

type CustomNodePart = {
  id: string
  profileId: string
  hardwareId: string
  readyHardwareId: string
  quantity: number
  deductionMm: number
  additionMm: number
}

type CustomNode = {
  id: string
  name: string
  formula: CustomNodeFormula
  quantityMode: CustomNodeQuantityMode
  quantityValue: number
  parts: CustomNodePart[]
}

type ReadyAssembly = {
  id: string
  article: string
  name: string
  sectionImageDataUrl?: string
  parts: CustomNodePart[]
}

type NodeTemplate = {
  id: string
  article: string
  name: string
  parts: CustomNodePart[]
}

type SystemNodeTemplateBinding = {
  frameSingleTrackAssembly: string
  frameDoubleTrackAssembly: string
  frameTripleTrackAssembly: string
  frameQuadrupleTrackAssembly: string
  fixFix: string
  fixSl: string
  slLeafBottomA: string
  slLeafBottomAFreeway: string
  slLeafPerimeterBcd: string
  slLeafVerticalBd: string
  slLeafTopC: string
  slSideToFrame: string
  slSideToFix: string
  slSideToSl: string
  slSideToPenal: string
  slSideToCorner: string
}

type NodeProfileBinding = {
  frameOuter65: string
  frameOuter80: string
  frameThermal: string
  frameCentral65: string
  frameCentral80: string
  fixFix: string
  fixSl: string
  slLeafBottomA: string
  slLeafPerimeterBcd: string
  slSideToFrame: string
  slSideToFix: string
  slSideToSl: string
  slSideToPenal: string
  slSideToCorner: string
}

type ProjectData = {
  projectNumber?: string
  preparedBy?: string
  widthMm: number
  heightMm: number
  frameSeries: FrameSeries
  leafBottomMode?: LeafBottomMode
  segmentsCount: number
  segments: SegmentType[]
  openingDirections: OpeningDirection[]
  segmentHeightOffset65Mm?: number
  segmentHeightOffset80Mm?: number
  wastePercent: number
  profiles: Profile[]
  hardware?: HardwareItem[]
  readyHardware?: ReadyHardwareItem[]
  readyAssemblies?: ReadyAssembly[]
  nodeTemplates?: NodeTemplate[]
  systemNodeTemplateBinding?: SystemNodeTemplateBinding
  nodeProfileBinding?: NodeProfileBinding
  customNodes?: CustomNode[]
}

const defaultProfiles: Profile[] = [
  { id: 'frame-outer-65', article: 'FO-65', name: 'Наружняя чаша 65', defaultLength: 6000 },
  { id: 'frame-thermal-65', article: 'TB-65', name: 'Термомост 65', defaultLength: 6000 },
  { id: 'frame-central-65', article: 'FC-65', name: 'Центральный профиль 65', defaultLength: 6000 },
  { id: 'frame-outer-80', article: 'FO-80', name: 'Наружняя чаша 80', defaultLength: 6000 },
  { id: 'frame-central-80', article: 'FC-80', name: 'Центральный профиль 80', defaultLength: 6000 },
  { id: 'frame-main', article: 'RM-70', name: 'Рама 70', defaultLength: 6000 },
  { id: 'fix-fix', article: 'FF-25', name: 'Импост FF-25', defaultLength: 6000 },
  { id: 'fix-sl', article: 'FS-40', name: 'Импост FS-40', defaultLength: 6000 },
  { id: 'sl-perimeter-a', article: 'SLA-01', name: 'SL оклейка A (низ)', defaultLength: 2800 },
  { id: 'sl-perimeter-b', article: 'SLB-02', name: 'SL оклейка B/C/D', defaultLength: 2800 },
  { id: 'sl-to-frame', article: 'SLR-18', name: 'SL рабочая сторона к раме', defaultLength: 6000 },
  { id: 'sl-to-fix', article: 'SLF-19', name: 'SL рабочая сторона к FIX', defaultLength: 6000 },
  { id: 'sl-to-sl', article: 'SLS-20', name: 'SL рабочая сторона к SL', defaultLength: 6000 },
  { id: 'sl-to-penal', article: 'SLP-21', name: 'SL рабочая сторона к пеналу', defaultLength: 6000 },
  { id: 'sl-to-corner', article: 'SLC-22', name: 'SL рабочая сторона к углу', defaultLength: 6000 },
]

const defaultNodeProfileBinding: NodeProfileBinding = {
  frameOuter65: 'frame-outer-65',
  frameOuter80: 'frame-outer-80',
  frameThermal: 'frame-thermal-65',
  frameCentral65: 'frame-central-65',
  frameCentral80: 'frame-central-80',
  fixFix: 'fix-fix',
  fixSl: 'fix-sl',
  slLeafBottomA: 'sl-perimeter-a',
  slLeafPerimeterBcd: 'sl-perimeter-b',
  slSideToFrame: 'sl-to-frame',
  slSideToFix: 'sl-to-fix',
  slSideToSl: 'sl-to-sl',
  slSideToPenal: 'sl-to-penal',
  slSideToCorner: 'sl-to-corner',
}

const defaultSystemNodeTemplateBinding: SystemNodeTemplateBinding = {
  frameSingleTrackAssembly: '',
  frameDoubleTrackAssembly: '',
  frameTripleTrackAssembly: '',
  frameQuadrupleTrackAssembly: '',
  fixFix: '',
  fixSl: '',
  slLeafBottomA: '',
  slLeafBottomAFreeway: '',
  slLeafPerimeterBcd: '',
  slLeafVerticalBd: '',
  slLeafTopC: '',
  slSideToFrame: '',
  slSideToFix: '',
  slSideToSl: '',
  slSideToPenal: '',
  slSideToCorner: '',
}

const defaultSlNodeTemplateNames = ['SL', 'SL FW', 'SL Motor', 'SL FW Motor'] as const

const getInitialProjectData = (): ProjectData | null => {
  return null
}

const getInitialActiveTab = (): ActiveTab => {
  if (typeof window !== 'undefined') {
    const hash = window.location.hash.replace('#', '')
    if (
      hash === 'config' ||
      hash === 'profiles' ||
      hash === 'hardware' ||
      hash === 'ready-hardware' ||
      hash === 'assemblies' ||
      hash === 'nodes' ||
      hash === 'settings'
    ) {
      return hash
    }
  }
  return 'config'
}

const normalizeHardware = (items: unknown): HardwareItem[] => {
  if (!Array.isArray(items)) return []
  return items
    .map((raw, index) => {
      const item = raw as Partial<HardwareItem>
      return {
        id: item.id || `hardware-${Date.now()}-${index}`,
        article: item.article || '',
        name: item.name || `Фурнитура ${index + 1}`,
      }
    })
    .filter((item) => item.article.trim() || item.name.trim())
}

const normalizeReadyHardware = (items: unknown): ReadyHardwareItem[] => {
  if (!Array.isArray(items)) return []
  return items
    .map((raw, index) => {
      const item = raw as Partial<ReadyHardwareItem>
      return {
        id: item.id || `ready-hardware-${Date.now()}-${index}`,
        article: item.article || '',
        name: item.name || `Готовая фурнитура ${index + 1}`,
      }
    })
    .filter((item) => item.article.trim() || item.name.trim())
}

const normalizeCustomNodes = (nodes: unknown): CustomNode[] => {
  if (!Array.isArray(nodes)) return []
  return nodes
    .map((raw, index) => {
      const item = raw as Partial<CustomNode> & { profileId?: string }
      const partsFromOld =
        item.profileId && (!item.parts || item.parts.length === 0)
          ? [
              {
                id: `migrated-${index}`,
                profileId: item.profileId,
                hardwareId: '',
                readyHardwareId: '',
                quantity: 1,
                deductionMm: 0,
                additionMm: 0,
              },
            ]
          : []
      const nextParts = Array.isArray(item.parts) ? item.parts : partsFromOld
      return {
        id: item.id || `node-${Date.now()}-${index}`,
        name: item.name || `Узел ${index + 1}`,
        formula: item.formula || 'framePerimeter',
        quantityMode: item.quantityMode || 'fixed',
        quantityValue: Math.max(1, Number(item.quantityValue) || 1),
        parts: nextParts
          .map((p, partIndex) => ({
            id: p.id || `${item.id || `node-${index}`}-part-${partIndex}`,
            profileId: p.profileId || '',
            hardwareId: (p as { hardwareId?: string }).hardwareId || '',
            readyHardwareId: (p as { readyHardwareId?: string }).readyHardwareId || '',
            quantity: Math.max(1, Number((p as { quantity?: number }).quantity ?? 1)),
            deductionMm: Math.max(0, Number((p as { deductionMm?: number }).deductionMm ?? 0)),
            additionMm: Math.max(0, Number((p as { additionMm?: number }).additionMm ?? 0)),
          }))
          .filter((p) => p.profileId || p.hardwareId || p.readyHardwareId),
      }
    })
    .filter((n) => n.parts.length > 0)
}

const normalizeReadyAssemblies = (assemblies: unknown): ReadyAssembly[] => {
  if (!Array.isArray(assemblies)) return []
  return assemblies
    .map((raw, index) => {
      const item = raw as Partial<ReadyAssembly>
      return {
        id: item.id || `assembly-${Date.now()}-${index}`,
        article: item.article || '',
        name: item.name || `Сборка ${index + 1}`,
        sectionImageDataUrl: item.sectionImageDataUrl,
        parts: (Array.isArray(item.parts) ? item.parts : [])
          .map((part, partIndex) => ({
            id: part.id || `${item.id || `assembly-${index}`}-part-${partIndex}`,
            profileId: part.profileId || '',
            hardwareId: (part as { hardwareId?: string }).hardwareId || '',
            readyHardwareId: (part as { readyHardwareId?: string }).readyHardwareId || '',
            quantity: Math.max(1, Number((part as { quantity?: number }).quantity ?? 1)),
            deductionMm: Math.max(0, Number((part as { deductionMm?: number }).deductionMm ?? 0)),
            additionMm: Math.max(0, Number((part as { additionMm?: number }).additionMm ?? 0)),
          }))
          .filter((part) => part.profileId || part.hardwareId || part.readyHardwareId),
      }
    })
    .filter((item) => item.parts.length > 0)
}

const normalizeNodeTemplates = (templates: unknown): NodeTemplate[] => {
  if (!Array.isArray(templates)) return []
  return templates
    .map((raw, index) => {
      const item = raw as Partial<NodeTemplate>
      return {
        id: item.id || `node-template-${Date.now()}-${index}`,
        article: item.article || '',
        name: item.name || `Узел ${index + 1}`,
        parts: (Array.isArray(item.parts) ? item.parts : [])
          .map((part, partIndex) => ({
            id: part.id || `${item.id || `node-template-${index}`}-part-${partIndex}`,
            profileId: part.profileId || '',
            hardwareId: (part as { hardwareId?: string }).hardwareId || '',
            readyHardwareId: (part as { readyHardwareId?: string }).readyHardwareId || '',
            quantity: Math.max(1, Number((part as { quantity?: number }).quantity ?? 1)),
            deductionMm: Math.max(0, Number((part as { deductionMm?: number }).deductionMm ?? 0)),
            additionMm: Math.max(0, Number((part as { additionMm?: number }).additionMm ?? 0)),
          }))
          .filter((part) => part.profileId || part.hardwareId || part.readyHardwareId),
      }
    })
    .filter((item) => item.parts.length > 0)
}

function App() {
  const [initialProject] = useState<ProjectData | null>(() => getInitialProjectData())
  const initialProfileDraft = { article: '', name: '', length: 3000 }

  const [activeTab, setActiveTab] = useState<ActiveTab>(() => getInitialActiveTab())
  const [projectNumber, setProjectNumber] = useState(initialProject?.projectNumber ?? '')
  const [preparedBy, setPreparedBy] = useState(initialProject?.preparedBy ?? '')
  const [widthMm, setWidthMm] = useState(initialProject?.widthMm ?? 6000)
  const [heightMm, setHeightMm] = useState(initialProject?.heightMm ?? 3000)
  const [frameSeries, setFrameSeries] = useState<FrameSeries>(initialProject?.frameSeries ?? '65')
  const [leafBottomMode, setLeafBottomMode] = useState<LeafBottomMode>(
    initialProject?.leafBottomMode ?? 'standard',
  )
  const [segmentsCount, setSegmentsCount] = useState(Math.max(1, initialProject?.segmentsCount ?? 3))
  const [segments, setSegments] = useState<SegmentType[]>(initialProject?.segments ?? ['FIX', 'FIX', 'SL'])
  const [openingDirections, setOpeningDirections] = useState<OpeningDirection[]>(
    initialProject?.openingDirections ?? ['right', 'right', 'right'],
  )
  const [profiles, setProfiles] = useState<Profile[]>(
    initialProject?.profiles?.length ? initialProject.profiles : defaultProfiles,
  )
  const [hardware, setHardware] = useState<HardwareItem[]>(normalizeHardware(initialProject?.hardware))
  const [newHardwareArticle, setNewHardwareArticle] = useState('')
  const [newHardwareName, setNewHardwareName] = useState('')
  const [hardwareSearch, setHardwareSearch] = useState('')
  const [readyHardware, setReadyHardware] = useState<ReadyHardwareItem[]>(
    normalizeReadyHardware(initialProject?.readyHardware),
  )
  const [newReadyHardwareArticle, setNewReadyHardwareArticle] = useState('')
  const [newReadyHardwareName, setNewReadyHardwareName] = useState('')
  const [readyHardwareSearch, setReadyHardwareSearch] = useState('')
  const [nodeProfileBinding, setNodeProfileBinding] = useState<NodeProfileBinding>({
    ...defaultNodeProfileBinding,
    ...(initialProject?.nodeProfileBinding ?? {}),
  })
  const [systemNodeTemplateBinding, setSystemNodeTemplateBinding] = useState<SystemNodeTemplateBinding>({
    ...defaultSystemNodeTemplateBinding,
    ...(initialProject?.systemNodeTemplateBinding ?? {}),
  })
  const [customNodes, setCustomNodes] = useState<CustomNode[]>(
    normalizeCustomNodes(initialProject?.customNodes),
  )
  const [readyAssemblies, setReadyAssemblies] = useState<ReadyAssembly[]>(
    normalizeReadyAssemblies(initialProject?.readyAssemblies),
  )
  const [nodeTemplates, setNodeTemplates] = useState<NodeTemplate[]>(
    normalizeNodeTemplates(initialProject?.nodeTemplates),
  )
  const [wastePercent, setWastePercent] = useState(initialProject?.wastePercent ?? 7)
  const [newProfileName, setNewProfileName] = useState(initialProfileDraft.name)
  const [newProfileArticle, setNewProfileArticle] = useState(initialProfileDraft.article)
  const [newProfileLength, setNewProfileLength] = useState(initialProfileDraft.length)
  const [newProfileSectionImageDataUrl, setNewProfileSectionImageDataUrl] = useState('')
  const [sectionPreviewUrl, setSectionPreviewUrl] = useState('')
  const [newNodeName, setNewNodeName] = useState('')
  const [newNodeFormula, setNewNodeFormula] = useState<CustomNodeFormula>('framePerimeter')
  const [newNodeQuantityMode, setNewNodeQuantityMode] = useState<CustomNodeQuantityMode>('fixed')
  const [newNodeQuantityValue, setNewNodeQuantityValue] = useState(1)
  const [newNodeParts, setNewNodeParts] = useState<CustomNodePart[]>([
    {
      id: 'new-part-1',
      profileId: '',
      hardwareId: '',
      readyHardwareId: '',
      quantity: 1,
      deductionMm: 0,
      additionMm: 0,
    },
  ])
  const [newAssemblyName, setNewAssemblyName] = useState('')
  const [newAssemblyArticle, setNewAssemblyArticle] = useState('')
  const [newAssemblySectionImageDataUrl, setNewAssemblySectionImageDataUrl] = useState('')
  const [newAssemblyParts, setNewAssemblyParts] = useState<CustomNodePart[]>([
    {
      id: 'new-assembly-part-1',
      profileId: '',
      hardwareId: '',
      readyHardwareId: '',
      quantity: 1,
      deductionMm: 0,
      additionMm: 0,
    },
  ])
  const [newNodeTemplateName, setNewNodeTemplateName] = useState('')
  const [newNodeTemplateArticle, setNewNodeTemplateArticle] = useState('')
  const [newNodeTemplateParts, setNewNodeTemplateParts] = useState<CustomNodePart[]>([
    {
      id: 'new-node-template-part-1',
      profileId: '',
      hardwareId: '',
      readyHardwareId: '',
      quantity: 1,
      deductionMm: 0,
      additionMm: 0,
    },
  ])
  const [serverSyncStatus, setServerSyncStatus] = useState('Нет синхронизации')
  const [serverLoadReady, setServerLoadReady] = useState(false)
  const [profileSearch, setProfileSearch] = useState('')
  const [assemblySearch, setAssemblySearch] = useState('')
  const [nodeTemplateSearch, setNodeTemplateSearch] = useState('')
  const [newAssemblyPartSearch, setNewAssemblyPartSearch] = useState<Record<string, string>>({})
  const [readyAssemblyPartSearch, setReadyAssemblyPartSearch] = useState<Record<string, string>>({})
  const [uiWarning, setUiWarning] = useState('')
  const [segmentHeightOffset65Mm, setSegmentHeightOffset65Mm] = useState(
    Math.max(0, initialProject?.segmentHeightOffset65Mm ?? 52),
  )
  const [segmentHeightOffset80Mm, setSegmentHeightOffset80Mm] = useState(
    Math.max(0, initialProject?.segmentHeightOffset80Mm ?? 67),
  )
  const warningTimerRef = useRef<number | undefined>(undefined)
  const lastWarningKeyRef = useRef('')
  const autosaveTimerRef = useRef<number | undefined>(undefined)

  const editableZeroNumber = (value: number) => (value === 0 ? '' : value)
  const showWarning = (message: string, key?: string) => {
    const warningKey = key ?? message
    if (lastWarningKeyRef.current === warningKey && uiWarning === message) return
    lastWarningKeyRef.current = warningKey
    setUiWarning(message)
    if (warningTimerRef.current) {
      window.clearTimeout(warningTimerRef.current)
    }
    warningTimerRef.current = window.setTimeout(() => {
      setUiWarning('')
      warningTimerRef.current = undefined
      lastWarningKeyRef.current = ''
    }, 30000)
  }

  const filteredProfiles = useMemo(() => {
    const q = profileSearch.trim().toLowerCase()
    if (!q) return profiles
    return profiles.filter(
      (p) => p.article.toLowerCase().includes(q) || p.name.toLowerCase().includes(q),
    )
  }, [profiles, profileSearch])

  const filteredReadyAssemblies = useMemo(() => {
    const q = assemblySearch.trim().toLowerCase()
    if (!q) return readyAssemblies
    return readyAssemblies.filter(
      (a) => a.article.toLowerCase().includes(q) || a.name.toLowerCase().includes(q),
    )
  }, [assemblySearch, readyAssemblies])

  const filteredNodeTemplates = useMemo(() => {
    const q = nodeTemplateSearch.trim().toLowerCase()
    if (!q) return nodeTemplates
    return nodeTemplates.filter(
      (n) => n.article.toLowerCase().includes(q) || n.name.toLowerCase().includes(q),
    )
  }, [nodeTemplateSearch, nodeTemplates])

  const filteredHardware = useMemo(() => {
    const q = hardwareSearch.trim().toLowerCase()
    if (!q) return hardware
    return hardware.filter(
      (h) => h.article.toLowerCase().includes(q) || h.name.toLowerCase().includes(q),
    )
  }, [hardware, hardwareSearch])

  const filteredReadyHardware = useMemo(() => {
    const q = readyHardwareSearch.trim().toLowerCase()
    if (!q) return readyHardware
    return readyHardware.filter(
      (h) => h.article.toLowerCase().includes(q) || h.name.toLowerCase().includes(q),
    )
  }, [readyHardware, readyHardwareSearch])

  const getProfilesByQuery = (query: string) => {
    const normalizedQuery = query.trim().toLowerCase()
    if (!normalizedQuery) return profiles
    return profiles.filter(
      (profile) =>
        profile.article.toLowerCase().includes(normalizedQuery) ||
        profile.name.toLowerCase().includes(normalizedQuery),
    )
  }

  const getHardwareByQuery = (query: string) => {
    const normalizedQuery = query.trim().toLowerCase()
    if (!normalizedQuery) return hardware
    return hardware.filter(
      (item) =>
        item.article.toLowerCase().includes(normalizedQuery) ||
        item.name.toLowerCase().includes(normalizedQuery),
    )
  }

  type AssemblyPartChoice =
    | { kind: 'profile'; id: string; label: string }
    | { kind: 'hardware'; id: string; label: string }
    | { kind: 'readyHardware'; id: string; label: string }

  const getAssemblyPartChoices = (query: string): AssemblyPartChoice[] => {
    const profileChoices = getProfilesByQuery(query).map((profile) => ({
      kind: 'profile' as const,
      id: profile.id,
      label: `${profile.article} | ${profile.name}`,
    }))
    const hardwareChoices = getHardwareByQuery(query).map((item) => ({
      kind: 'hardware' as const,
      id: item.id,
      label: `${item.article} | ${item.name}`,
    }))
    const readyHardwareChoices = readyHardware
      .filter((item) => {
        const normalizedQuery = query.trim().toLowerCase()
        if (!normalizedQuery) return true
        return (
          item.article.toLowerCase().includes(normalizedQuery) ||
          item.name.toLowerCase().includes(normalizedQuery)
        )
      })
      .map((item) => ({
        kind: 'readyHardware' as const,
        id: item.id,
        label: `${item.article} | ${item.name}`,
      }))
    return [...profileChoices, ...hardwareChoices, ...readyHardwareChoices]
  }

  const getAssemblyPartSelectedLabel = (part: CustomNodePart) => {
    if (part.profileId) {
      const profile = profiles.find((item) => item.id === part.profileId)
      if (profile) return `${profile.article} | ${profile.name}`
    }
    if (part.hardwareId) {
      const hw = hardware.find((item) => item.id === part.hardwareId)
      if (hw) return `${hw.article} | ${hw.name}`
    }
    if (part.readyHardwareId) {
      const kit = readyHardware.find((item) => item.id === part.readyHardwareId)
      if (kit) return `${kit.article} | ${kit.name}`
    }
    return ''
  }

  const resolveAssemblyPartSelection = (part: CustomNodePart, searchValue: string): CustomNodePart => {
    if (part.profileId || part.hardwareId || part.readyHardwareId) return part
    const query = searchValue.trim()
    if (!query) return part
    const allChoices = getAssemblyPartChoices('')
    const exact = allChoices.find((choice) => choice.label.toLowerCase() === query.toLowerCase())
    const filteredChoices = getAssemblyPartChoices(query)
    const matched = exact ?? (filteredChoices.length === 1 ? filteredChoices[0] : null)
    if (!matched) return part
    return {
      ...part,
      profileId: matched.kind === 'profile' ? matched.id : '',
      hardwareId: matched.kind === 'hardware' ? matched.id : '',
      readyHardwareId: matched.kind === 'readyHardware' ? matched.id : '',
    }
  }

  const isFreewayAssembly = (assembly: ReadyAssembly) => {
    const label = `${assembly.article} ${assembly.name}`.toLowerCase()
    return label.includes('freeway') || label.includes('fw')
  }

  const isOkleykaAssemblyAllowed = (assemblyId: string) => {
    if (!assemblyId) return true
    const assembly = readyAssemblies.find((item) => item.id === assemblyId)
    if (!assembly) return true
    const freewayAssembly = isFreewayAssembly(assembly)
    return leafBottomMode === 'freeway' ? freewayAssembly : !freewayAssembly
  }

  const normalizedSegments = useMemo(() => {
    return Array.from({ length: segmentsCount }, (_, index) => segments[index] ?? 'FIX')
  }, [segments, segmentsCount])

  useEffect(() => {
    if (frameSeries !== '65') {
      return
    }
    if (activeTab !== 'config') return
    const hasSliding = segments.some((segment) => segment === 'SL')
    if (!hasSliding) return
    setSegments((current) => current.map(() => 'FIX'))
    showWarning(
      'Серия 65 не поддерживает откатные створки. Все сегменты автоматически переведены в FIX.',
      'series65-auto-fix',
    )
  }, [frameSeries, segments, activeTab])

  const activeSlBottomBindingId =
    leafBottomMode === 'freeway'
      ? systemNodeTemplateBinding.slLeafBottomAFreeway
      : systemNodeTemplateBinding.slLeafBottomA
  const activeSlVerticalBdBindingId =
    systemNodeTemplateBinding.slLeafVerticalBd || systemNodeTemplateBinding.slLeafPerimeterBcd
  const activeSlTopCBindingId =
    systemNodeTemplateBinding.slLeafTopC || systemNodeTemplateBinding.slLeafPerimeterBcd

  useEffect(() => {
    if (activeTab !== 'config' && activeTab !== 'settings') return
    if (!normalizedSegments.some((segment) => segment === 'SL')) return
    if (activeSlBottomBindingId) return
    const modeLabel = leafBottomMode === 'freeway' ? 'FreeWay' : 'Обычный профиль'
    showWarning(
      `Для режима "${modeLabel}" не назначен узел/сборка "SL оклейка A (низ)" в Настройках.`,
      `missing-sl-bottom-${leafBottomMode}`,
    )
  }, [normalizedSegments, activeSlBottomBindingId, leafBottomMode, activeTab])

  useEffect(() => {
    if (activeTab !== 'config' && activeTab !== 'settings') return
    if (!normalizedSegments.some((segment) => segment === 'SL')) return
    const invalidBindings = [
      activeSlBottomBindingId,
      activeSlVerticalBdBindingId,
      activeSlTopCBindingId,
    ].filter((id) => id && !isOkleykaAssemblyAllowed(id))
    if (invalidBindings.length === 0) return
    const modeLabel = leafBottomMode === 'freeway' ? 'FreeWay' : 'Обычный профиль'
    showWarning(
      `Для режима "${modeLabel}" у оклейки выбраны неподходящие готовые профили. Проверьте пометку FW.`,
      `invalid-okleyka-fw-${leafBottomMode}`,
    )
  }, [
    normalizedSegments,
    activeTab,
    leafBottomMode,
    activeSlBottomBindingId,
    activeSlVerticalBdBindingId,
    activeSlTopCBindingId,
    readyAssemblies,
  ])

  useEffect(() => {
    return () => {
      if (warningTimerRef.current) {
        window.clearTimeout(warningTimerRef.current)
      }
    }
  }, [])

  const normalizedOpeningDirections = useMemo(() => {
    return Array.from(
      { length: segmentsCount },
      (_, index) => openingDirections[index] ?? 'right',
    )
  }, [openingDirections, segmentsCount])

  const panelWidth = useMemo(() => {
    if (segmentsCount === 0) return 0
    return widthMm / segmentsCount
  }, [widthMm, segmentsCount])
  const segmentHeightOffsetMm = frameSeries === '65' ? segmentHeightOffset65Mm : segmentHeightOffset80Mm
  const segmentHeightMm = useMemo(
    () => Math.max(0, heightMm - segmentHeightOffsetMm),
    [heightMm, segmentHeightOffsetMm],
  )

  const calculations = useMemo(() => {
    const perimeter = 2 * (widthMm + heightMm)
    const verticalJoints = Math.max(segmentsCount - 1, 0)
    let fixFixJoints = 0
    let fixSlJoints = 0

    for (let i = 0; i < normalizedSegments.length - 1; i += 1) {
      const left = normalizedSegments[i]
      const right = normalizedSegments[i + 1]
      if (left === 'FIX' && right === 'FIX') {
        fixFixJoints += 1
      } else {
        fixSlJoints += 1
      }
    }

    const slidingPanels = normalizedSegments.filter((item) => item === 'SL').length
    const slidingPerimeter = slidingPanels * 2 * (panelWidth + segmentHeightMm)

    return {
      perimeter,
      verticalJoints,
      fixFixJoints,
      fixSlJoints,
      slidingPanels,
      slidingPerimeter,
    }
  }, [heightMm, normalizedSegments, panelWidth, segmentHeightMm, segmentsCount, widthMm])

  const frameSystem = useMemo(() => {
    const slidingIndexes = normalizedSegments
      .map((segment, index) => (segment === 'SL' ? index : -1))
      .filter((index) => index >= 0)
    const slidingCount = slidingIndexes.length

    if (slidingCount === 0) {
      return {
        type: 'single-track' as FrameSystemType,
        title: 'Однотрековая',
        description: 'Только фиксированные сегменты',
      }
    }

    if (slidingCount === 1) {
      return {
        type: 'double-track' as FrameSystemType,
        title: 'Двухтрековая',
        description: 'Фиксы + одна подвижная',
      }
    }

    if (slidingCount === 2) {
      const [firstIndex, secondIndex] = slidingIndexes
      const firstDir = normalizedOpeningDirections[firstIndex]
      const secondDir = normalizedOpeningDirections[secondIndex]
      const fixedCount = normalizedSegments.filter((segment) => segment === 'FIX').length
      const isThreeTrackPattern = normalizedSegments.length === 3 && fixedCount === 1
      if (firstDir === secondDir && isThreeTrackPattern) {
        return {
          type: 'triple-track' as FrameSystemType,
          title: 'Трехтрековая',
          description: '2 SL в одну сторону + 1 FIX',
        }
      }
      if (firstDir === secondDir) {
        return {
          type: 'triple-track' as FrameSystemType,
          title: 'Трехтрековая',
          description: 'Эвристика: 2 SL в одну сторону (вне базового шаблона)',
        }
      }
      return {
        type: 'double-track' as FrameSystemType,
        title: 'Двухтрековая',
        description: 'Две подвижные в разные стороны',
      }
    }

    if (slidingCount === 3) {
      const slidingDirections = slidingIndexes.map((index) => normalizedOpeningDirections[index])
      const allSameDirection = slidingDirections.every((dir) => dir === slidingDirections[0])
      const fixedCount = normalizedSegments.filter((segment) => segment === 'FIX').length
      const isFourTrackPattern = normalizedSegments.length === 5 && fixedCount === 2
      if (allSameDirection && isFourTrackPattern) {
        return {
          type: 'quadruple-track' as FrameSystemType,
          title: 'Четырехтрековая',
          description: '3 SL в одну сторону + 2 FIX',
        }
      }
      if (allSameDirection) {
        return {
          type: 'quadruple-track' as FrameSystemType,
          title: 'Четырехтрековая',
          description: 'Эвристика: 3 SL в одну сторону (вне базового шаблона)',
        }
      }
      const leftCount = slidingDirections.filter((dir) => dir === 'left').length
      const rightCount = slidingDirections.filter((dir) => dir === 'right').length
      const hasDominantDirection = leftCount === 2 || rightCount === 2
      if (hasDominantDirection) {
        return {
          type: 'triple-track' as FrameSystemType,
          title: 'Трехтрековая',
          description: 'Шаблон: 3 SL со смешанным открыванием (2 в одну сторону, 1 в другую)',
        }
      }
      return {
        type: null,
        title: 'Не определено',
        description: 'Не удалось определить систему: 3 SL с разнонаправленным открыванием',
      }
    }

    if (slidingCount === 4) {
      const isFiveBlockPattern = normalizedSegments.length === 7
      const matchesGroupedLayout =
        normalizedSegments[0] === 'FIX' &&
        normalizedSegments[1] === 'SL' &&
        normalizedSegments[2] === 'SL' &&
        normalizedSegments[3] === 'FIX' &&
        normalizedSegments[4] === 'SL' &&
        normalizedSegments[5] === 'SL' &&
        normalizedSegments[6] === 'FIX'
      if (isFiveBlockPattern && matchesGroupedLayout) {
        return {
          type: 'triple-track' as FrameSystemType,
          title: 'Трехтрековая',
          description: 'Шаблон: FIX + 2SL + FIX + 2SL + FIX',
        }
      }

      const slidingDirections = slidingIndexes.map((index) => normalizedOpeningDirections[index])
      const leftCount = slidingDirections.filter((dir) => dir === 'left').length
      const rightCount = slidingDirections.filter((dir) => dir === 'right').length
      if (leftCount === 2 && rightCount === 2) {
        return {
          type: 'triple-track' as FrameSystemType,
          title: 'Трехтрековая',
          description: 'Шаблон: 4 SL (2 влево + 2 вправо)',
        }
      }
      if (leftCount === 4 || rightCount === 4) {
        return {
          type: 'quadruple-track' as FrameSystemType,
          title: 'Четырехтрековая',
          description: 'Эвристика: 4 SL в одну сторону',
        }
      }
      return {
        type: null,
        title: 'Не определено',
        description: 'Не удалось определить систему: 4 SL с произвольным направлением',
      }
    }

    if (slidingCount === 5) {
      const slidingDirections = slidingIndexes.map((index) => normalizedOpeningDirections[index])
      const leftCount = slidingDirections.filter((dir) => dir === 'left').length
      const rightCount = slidingDirections.filter((dir) => dir === 'right').length
      const isSameDirection = leftCount === 5 || rightCount === 5
      const isDominantDirection = leftCount === 3 || rightCount === 3
      if (isSameDirection || isDominantDirection) {
        return {
          type: 'quadruple-track' as FrameSystemType,
          title: 'Четырехтрековая',
          description: 'Шаблон: 5 SL (5 в одну сторону или 3+2 по направлениям)',
        }
      }
      return {
        type: null,
        title: 'Не определено',
        description: 'Не удалось определить систему: 5 SL с нестандартным направлением',
      }
    }

    return {
      type: null,
      title: 'Не определено',
      description: 'Для этой конфигурации пока нет шаблона системы рамы',
    }
  }, [normalizedOpeningDirections, normalizedSegments])

  const frameComposition = useMemo(() => {
    if (!frameSystem.type) return []

    const getProfileById = (id: string, fallbackName: string) => {
      const profile = profiles.find((item) => item.id === id)
      if (!profile) return null
      return {
        article: profile.article,
        name: profile.name || fallbackName,
        stockLengthMm: profile.defaultLength || 6000,
        sectionImageDataUrl: profile.sectionImageDataUrl,
      }
    }

    const frameOuter65 = getProfileById(nodeProfileBinding.frameOuter65, 'Наружняя чаша 65')
    const frameOuterTop = getProfileById(
      frameSeries === '80' ? nodeProfileBinding.frameOuter80 : nodeProfileBinding.frameOuter65,
      `Наружняя чаша ${frameSeries}`,
    )
    const frameThermal65 = getProfileById(nodeProfileBinding.frameThermal, 'Термомост 65')
    const frameThermalTop = getProfileById(nodeProfileBinding.frameThermal, 'Термомост 65')
    const frameCentral65 = getProfileById(nodeProfileBinding.frameCentral65, 'Центральный профиль 65')
    const frameCentralTop = getProfileById(
      frameSeries === '80' ? nodeProfileBinding.frameCentral80 : nodeProfileBinding.frameCentral65,
      `Центральный профиль ${frameSeries}`,
    )

    const sidesAndBottomLength = widthMm + 2 * heightMm
    const topLength = widthMm

    const splitByTopAndOther = (
      part: string,
      count: number,
      profile65: { article: string; name: string; stockLengthMm: number; sectionImageDataUrl?: string } | null,
      profileTop: { article: string; name: string; stockLengthMm: number; sectionImageDataUrl?: string } | null,
    ) =>
      [
        profile65
          ? {
              part: `${part} (бока+низ)`,
              count,
              article: profile65.article,
              profileName: profile65.name,
              stockLengthMm: profile65.stockLengthMm,
              sectionImageDataUrl: profile65.sectionImageDataUrl,
              lengthPerPartMm: sidesAndBottomLength,
              totalLengthMm: sidesAndBottomLength * count,
            }
          : null,
        profileTop
          ? {
              part: `${part} (верх)`,
              count,
              article: profileTop.article,
              profileName: profileTop.name,
              stockLengthMm: profileTop.stockLengthMm,
              sectionImageDataUrl: profileTop.sectionImageDataUrl,
              lengthPerPartMm: topLength,
              totalLengthMm: topLength * count,
            }
          : null,
      ].filter(Boolean) as {
        part: string
        count: number
        article: string
        profileName: string
        stockLengthMm: number
        sectionImageDataUrl?: string
        lengthPerPartMm: number
        totalLengthMm: number
      }[]

    const templates: Record<FrameSystemType, { part: string; count: number }[]> = {
      'single-track': [
        { part: 'Наружняя чаша', count: 2 },
        { part: 'Термомост', count: 1 },
      ],
      'double-track': [
        { part: 'Наружняя чаша', count: 2 },
        { part: 'Термомост', count: 2 },
        { part: 'Центральный профиль рамы', count: 1 },
      ],
      'triple-track': [
        { part: 'Наружняя чаша', count: 2 },
        { part: 'Термомост', count: 3 },
        { part: 'Центральный профиль рамы', count: 2 },
      ],
      'quadruple-track': [
        { part: 'Наружняя чаша', count: 2 },
        { part: 'Термомост', count: 4 },
        { part: 'Центральный профиль рамы', count: 3 },
      ],
    }

    return templates[frameSystem.type].flatMap((item) => {
      if (item.part === 'Наружняя чаша') {
        return splitByTopAndOther(item.part, item.count, frameOuter65, frameOuterTop)
      }
      if (item.part === 'Термомост') {
        return splitByTopAndOther(item.part, item.count, frameThermal65, frameThermalTop)
      }
      return splitByTopAndOther(item.part, item.count, frameCentral65, frameCentralTop)
    })
  }, [frameSeries, frameSystem.type, nodeProfileBinding, profiles, widthMm, heightMm])

  const selectedFrameAssembly = useMemo(() => {
    if (!frameSystem.type) return null
    const frameAssemblyBySystem: Record<FrameSystemType, string> = {
      'single-track': systemNodeTemplateBinding.frameSingleTrackAssembly,
      'double-track': systemNodeTemplateBinding.frameDoubleTrackAssembly,
      'triple-track': systemNodeTemplateBinding.frameTripleTrackAssembly,
      'quadruple-track': systemNodeTemplateBinding.frameQuadrupleTrackAssembly,
    }
    const assemblyId = frameAssemblyBySystem[frameSystem.type]
    if (!assemblyId) return null
    return readyAssemblies.find((item) => item.id === assemblyId) ?? null
  }, [frameSystem.type, readyAssemblies, systemNodeTemplateBinding])

  const selectedSlBottomAssembly = useMemo(() => {
    const assemblyId =
      leafBottomMode === 'freeway'
        ? systemNodeTemplateBinding.slLeafBottomAFreeway
        : systemNodeTemplateBinding.slLeafBottomA
    if (!assemblyId) return null
    return readyAssemblies.find((item) => item.id === assemblyId) ?? null
  }, [
    readyAssemblies,
    leafBottomMode,
    systemNodeTemplateBinding.slLeafBottomA,
    systemNodeTemplateBinding.slLeafBottomAFreeway,
  ])

  const selectedSlVerticalBdAssembly = useMemo(() => {
    const assemblyId =
      systemNodeTemplateBinding.slLeafVerticalBd || systemNodeTemplateBinding.slLeafPerimeterBcd
    if (!assemblyId) return null
    return readyAssemblies.find((item) => item.id === assemblyId) ?? null
  }, [
    readyAssemblies,
    systemNodeTemplateBinding.slLeafVerticalBd,
    systemNodeTemplateBinding.slLeafPerimeterBcd,
  ])

  const selectedSlTopCAssembly = useMemo(() => {
    const assemblyId = systemNodeTemplateBinding.slLeafTopC || systemNodeTemplateBinding.slLeafPerimeterBcd
    if (!assemblyId) return null
    return readyAssemblies.find((item) => item.id === assemblyId) ?? null
  }, [
    readyAssemblies,
    systemNodeTemplateBinding.slLeafTopC,
    systemNodeTemplateBinding.slLeafPerimeterBcd,
  ])

  const readyAssemblyCompositionRows = useMemo(() => {
    const rows: {
      key: string
      zone: string
      article: string
      name: string
      quantity: number
      deductionLabel: string
      lengthPerUnitMm: number | null
      totalLengthMm: number | null
    }[] = []

    const addAssemblyRow = (
      key: string,
      zone: string,
      assembly: ReadyAssembly | null,
      quantity: number,
      baseLengthPerUnitMm: number,
    ) => {
      if (!assembly || quantity <= 0) return
      if (!isOkleykaAssemblyAllowed(assembly.id)) return
      const adjustments = [...new Set(assembly.parts.map((part) => part.deductionMm - part.additionMm))]
      const hasSingleAdjustment = adjustments.length === 1
      const adjustment = hasSingleAdjustment ? adjustments[0] : 0
      const unitLength = hasSingleAdjustment ? Math.max(baseLengthPerUnitMm - adjustment, 0) : null
      rows.push({
        key,
        zone,
        article: assembly.article,
        name: assembly.name,
        quantity,
        deductionLabel: hasSingleAdjustment ? String(adjustment) : 'разные',
        lengthPerUnitMm: unitLength,
        totalLengthMm: unitLength === null ? null : quantity * unitLength,
      })
    }

    if (selectedFrameAssembly) {
      const deductions = [...new Set(selectedFrameAssembly.parts.map((part) => part.deductionMm))]
      const hasSingleDeduction = deductions.length === 1
      const deduction = hasSingleDeduction ? deductions[0] : 0
      rows.push({
        key: `frame-${selectedFrameAssembly.id}`,
        zone: `Рама (${frameSystem.title})`,
        article: selectedFrameAssembly.article,
        name: selectedFrameAssembly.name,
        quantity: 1,
        deductionLabel: hasSingleDeduction ? String(deduction) : 'разные',
        lengthPerUnitMm: hasSingleDeduction ? Math.max(calculations.perimeter - deduction, 0) : null,
        totalLengthMm: hasSingleDeduction ? Math.max(calculations.perimeter - deduction, 0) : null,
      })
    }

    addAssemblyRow(
      `sl-bottom-${selectedSlBottomAssembly?.id ?? 'none'}`,
      'SL оклейка A (низ)',
      selectedSlBottomAssembly,
      calculations.slidingPanels,
      panelWidth,
    )
    addAssemblyRow(
      `sl-bd-${selectedSlVerticalBdAssembly?.id ?? 'none'}`,
      'SL оклейка B/D (вертикали)',
      selectedSlVerticalBdAssembly,
      calculations.slidingPanels * 2,
      segmentHeightMm,
    )
    addAssemblyRow(
      `sl-c-${selectedSlTopCAssembly?.id ?? 'none'}`,
      'SL оклейка C (верх)',
      selectedSlTopCAssembly,
      calculations.slidingPanels,
      panelWidth,
    )

    return rows
  }, [
    calculations.perimeter,
    calculations.slidingPanels,
    frameSystem.title,
    panelWidth,
    segmentHeightMm,
    selectedFrameAssembly,
    selectedSlBottomAssembly,
    selectedSlVerticalBdAssembly,
    selectedSlTopCAssembly,
    leafBottomMode,
  ])

  const customNodeSpecItems = useMemo(() => {
    const getFormulaLength = (formula: CustomNodeFormula) => {
      switch (formula) {
        case 'framePerimeter':
          return calculations.perimeter
        case 'frameTop':
          return widthMm
        case 'frameSidesBottom':
          return widthMm + 2 * heightMm
        case 'frameHeight':
          return segmentHeightMm
        case 'panelWidth':
          return panelWidth
        case 'slPerimeterSingle':
          return 2 * (panelWidth + segmentHeightMm)
        default:
          return 0
      }
    }

    const getQuantity = (mode: CustomNodeQuantityMode, value: number) => {
      switch (mode) {
        case 'fixed':
          return value
        case 'slidingCount':
          return calculations.slidingPanels
        case 'segmentCount':
          return segmentsCount
        case 'verticalJoints':
          return calculations.verticalJoints
        default:
          return 0
      }
    }

    return customNodes
      .flatMap((node) => {
        const baseLength = getFormulaLength(node.formula)
        const qty = getQuantity(node.quantityMode, node.quantityValue)
        return node.parts
          .map((part) => {
            if (part.hardwareId) {
              const hw = hardware.find((h) => h.id === part.hardwareId)
              if (!hw) return null
              const pieceCount = part.quantity * qty
              if (pieceCount <= 0) return null
              return {
                key: `custom-${node.id}-${part.id}-hw`,
                article: hw.article,
                profile: hw.name,
                zone: `Пользовательский узел: ${node.name}`,
                partQuantity: part.quantity,
                deductionMm: 0,
                additionMm: 0,
                totalLengthMm: pieceCount,
                stockLengthMm: 1,
                isHardware: true,
              }
            }
            if (part.readyHardwareId) {
              const kit = readyHardware.find((item) => item.id === part.readyHardwareId)
              if (!kit) return null
              const pieceCount = part.quantity * qty
              if (pieceCount <= 0) return null
              return {
                key: `custom-${node.id}-${part.id}-rhw`,
                article: kit.article,
                profile: kit.name,
                zone: `Пользовательский узел: ${node.name}`,
                partQuantity: part.quantity,
                deductionMm: 0,
                additionMm: 0,
                totalLengthMm: pieceCount,
                stockLengthMm: 1,
                isHardware: true,
              }
            }
            const profile = profiles.find((p) => p.id === part.profileId)
            if (!profile) return null
            return {
              key: `custom-${node.id}-${part.id}`,
              article: profile.article,
              profile: profile.name,
              zone: `Пользовательский узел: ${node.name}`,
              partQuantity: part.quantity,
              deductionMm: part.deductionMm,
              additionMm: part.additionMm,
              totalLengthMm: Math.max(baseLength - part.deductionMm + part.additionMm, 0) * qty * part.quantity,
              stockLengthMm: profile.defaultLength,
              sectionImageDataUrl: profile.sectionImageDataUrl,
            }
          })
          .filter(Boolean)
      })
      .filter((item) => (item?.totalLengthMm ?? 0) > 0) as Omit<
      SpecItem,
      'barsCount' | 'withWasteLengthMm'
    >[]
  }, [
    calculations,
    customNodes,
    hardware,
    readyHardware,
    heightMm,
    panelWidth,
    profiles,
    segmentHeightMm,
    segmentsCount,
    widthMm,
  ])

  const spec = useMemo<SpecItem[]>(() => {
    const getProfileById = (
      id: string,
      fallbackName: string,
      fallbackLength = 6000,
      fallbackArticle = '',
    ) => {
      const profile = profiles.find((item) => item.id === id)
      return {
        article: profile?.article ?? fallbackArticle,
        name: profile?.name ?? fallbackName,
        stockLengthMm: profile?.defaultLength ?? fallbackLength,
        sectionImageDataUrl: profile?.sectionImageDataUrl,
      }
    }
    const getNodeTemplateById = (id: string) => nodeTemplates.find((item) => item.id === id)
    const getReadyAssemblyById = (id: string) => readyAssemblies.find((item) => item.id === id)

    const assemblyLineHardwareInstances = (keyPrefix: string, totalLengthMm: number): number => {
      if (totalLengthMm <= 0) return 0
      const pw = panelWidth
      const sh = segmentHeightMm
      if (keyPrefix === 'frame-system-assembly') return 1
      if (keyPrefix === 'sl-leaf-bottom-a' || keyPrefix === 'sl-leaf-perimeter-c') {
        return pw > 0 ? totalLengthMm / pw : 0
      }
      if (keyPrefix === 'sl-leaf-perimeter-bd') {
        const d = 2 * sh
        return d > 0 ? totalLengthMm / d : 0
      }
      if (keyPrefix === 'fix-fix' || keyPrefix === 'fix-sl') {
        return sh > 0 ? totalLengthMm / sh : 0
      }
      if (keyPrefix.startsWith('sl-side-')) {
        return sh > 0 ? totalLengthMm / sh : 0
      }
      return totalLengthMm > 0 ? 1 : 0
    }

    const expandWithReadyAssembly = (
      keyPrefix: string,
      zone: string,
      totalLengthMm: number,
      assemblyId: string,
    ): Omit<SpecItem, 'barsCount' | 'withWasteLengthMm'>[] => {
      const assembly = getReadyAssemblyById(assemblyId)
      if (!assembly || totalLengthMm <= 0) return []
      const instances = assemblyLineHardwareInstances(keyPrefix, totalLengthMm)
      return assembly.parts
        .map((part) => {
          if (part.hardwareId) {
            const hwItem = hardware.find((item) => item.id === part.hardwareId)
            if (!hwItem) return null
            const pieceCount = Math.max(0, Math.ceil(part.quantity * instances - 1e-9))
            if (pieceCount <= 0) return null
            return {
              key: `${keyPrefix}-${assembly.id}-${part.id}-hw`,
              article: hwItem.article,
              profile: hwItem.name,
              zone: `${zone} (${assembly.article ? `${assembly.article} | ` : ''}${assembly.name})`,
              partQuantity: part.quantity,
              deductionMm: part.deductionMm,
              additionMm: part.additionMm,
              totalLengthMm: pieceCount,
              stockLengthMm: 1,
              isHardware: true,
            }
          }
          if (part.readyHardwareId) {
            const kit = readyHardware.find((item) => item.id === part.readyHardwareId)
            if (!kit) return null
            const pieceCount = Math.max(0, Math.ceil(part.quantity * instances - 1e-9))
            if (pieceCount <= 0) return null
            return {
              key: `${keyPrefix}-${assembly.id}-${part.id}-rhw`,
              article: kit.article,
              profile: kit.name,
              zone: `${zone} (${assembly.article ? `${assembly.article} | ` : ''}${assembly.name})`,
              partQuantity: part.quantity,
              deductionMm: part.deductionMm,
              additionMm: part.additionMm,
              totalLengthMm: pieceCount,
              stockLengthMm: 1,
              isHardware: true,
            }
          }
          const profile = profiles.find((item) => item.id === part.profileId)
          if (!profile) return null
          return {
            key: `${keyPrefix}-${assembly.id}-${part.id}`,
            article: profile.article,
            profile: profile.name,
            zone: `${zone} (${assembly.article ? `${assembly.article} | ` : ''}${assembly.name})`,
            partQuantity: part.quantity,
            deductionMm: part.deductionMm,
            additionMm: part.additionMm,
            totalLengthMm: Math.max(totalLengthMm - part.deductionMm + part.additionMm, 0) * part.quantity,
            stockLengthMm: profile.defaultLength,
            sectionImageDataUrl: profile.sectionImageDataUrl,
          }
        })
        .filter(Boolean) as Omit<SpecItem, 'barsCount' | 'withWasteLengthMm'>[]
    }
    const expandWithNodeTemplate = (
      keyPrefix: string,
      zone: string,
      totalLengthMm: number,
      templateId: string,
      profileId: string,
      fallbackName: string,
    ): Omit<SpecItem, 'barsCount' | 'withWasteLengthMm'>[] => {
      const nodeTemplate = getNodeTemplateById(templateId)
      if (nodeTemplate && nodeTemplate.parts.length > 0) {
        const instances = assemblyLineHardwareInstances(keyPrefix, totalLengthMm)
        return nodeTemplate.parts
          .map((part) => {
            if (part.hardwareId) {
              const hwItem = hardware.find((item) => item.id === part.hardwareId)
              if (!hwItem) return null
              const pieceCount = Math.max(0, Math.ceil(part.quantity * instances - 1e-9))
              if (pieceCount <= 0) return null
              return {
                key: `${keyPrefix}-${nodeTemplate.id}-${part.id}-hw`,
                article: hwItem.article,
                profile: hwItem.name,
                zone: `${zone} (узел: ${nodeTemplate.article ? `${nodeTemplate.article} | ` : ''}${nodeTemplate.name})`,
                partQuantity: part.quantity,
                deductionMm: part.deductionMm,
                additionMm: part.additionMm,
                totalLengthMm: pieceCount,
                stockLengthMm: 1,
                isHardware: true,
              }
            }
            if (part.readyHardwareId) {
              const kit = readyHardware.find((item) => item.id === part.readyHardwareId)
              if (!kit) return null
              const pieceCount = Math.max(0, Math.ceil(part.quantity * instances - 1e-9))
              if (pieceCount <= 0) return null
              return {
                key: `${keyPrefix}-${nodeTemplate.id}-${part.id}-rhw`,
                article: kit.article,
                profile: kit.name,
                zone: `${zone} (узел: ${nodeTemplate.article ? `${nodeTemplate.article} | ` : ''}${nodeTemplate.name})`,
                partQuantity: part.quantity,
                deductionMm: part.deductionMm,
                additionMm: part.additionMm,
                totalLengthMm: pieceCount,
                stockLengthMm: 1,
                isHardware: true,
              }
            }
            const profile = profiles.find((item) => item.id === part.profileId)
            if (!profile) return null
            return {
              key: `${keyPrefix}-${nodeTemplate.id}-${part.id}`,
              article: profile.article,
              profile: profile.name,
              zone: `${zone} (узел: ${nodeTemplate.article ? `${nodeTemplate.article} | ` : ''}${nodeTemplate.name})`,
              partQuantity: part.quantity,
              deductionMm: part.deductionMm,
              additionMm: part.additionMm,
              totalLengthMm: Math.max(totalLengthMm - part.deductionMm + part.additionMm, 0) * part.quantity,
              stockLengthMm: profile.defaultLength,
              sectionImageDataUrl: profile.sectionImageDataUrl,
            }
          })
          .filter(Boolean) as Omit<SpecItem, 'barsCount' | 'withWasteLengthMm'>[]
      }
      const profile = getProfileById(profileId, fallbackName)
      return [
        {
          key: keyPrefix,
          article: profile.article,
          profile: profile.name,
          zone,
          partQuantity: 1,
          deductionMm: 0,
          additionMm: 0,
          totalLengthMm,
          stockLengthMm: profile.stockLengthMm,
          sectionImageDataUrl: profile.sectionImageDataUrl,
        },
      ]
    }
    const expandWithTemplateOrAssembly = (
      keyPrefix: string,
      zone: string,
      totalLengthMm: number,
      templateOrAssemblyId: string,
      profileId: string,
      fallbackName: string,
      enforceOkleykaAssemblyMode = false,
    ): Omit<SpecItem, 'barsCount' | 'withWasteLengthMm'>[] => {
      if (templateOrAssemblyId) {
        const fromAssembly = expandWithReadyAssembly(
          keyPrefix,
          zone,
          totalLengthMm,
          templateOrAssemblyId,
        )
        if (fromAssembly.length > 0) {
          if (enforceOkleykaAssemblyMode && !isOkleykaAssemblyAllowed(templateOrAssemblyId)) {
            return []
          }
          return fromAssembly
        }

        const hasTemplate = getNodeTemplateById(templateOrAssemblyId)
        if (hasTemplate) {
          return expandWithNodeTemplate(
            keyPrefix,
            zone,
            totalLengthMm,
            templateOrAssemblyId,
            profileId,
            fallbackName,
          )
        }

        // Если в настройках выбран id, но он не найден (устаревшие данные),
        // не подставляем fallback "1 шт", чтобы не искажать спецификацию.
        return []
      }

      return expandWithNodeTemplate(
        keyPrefix,
        zone,
        totalLengthMm,
        '',
        profileId,
        fallbackName,
      )
    }

    const fixFix = getProfileById(nodeProfileBinding.fixFix, 'Импост FIX/FIX')
    const fixSl = getProfileById(nodeProfileBinding.fixSl, 'Импост FIX/SL')
    const selectedSlBottomBindingId =
      leafBottomMode === 'freeway'
        ? systemNodeTemplateBinding.slLeafBottomAFreeway
        : systemNodeTemplateBinding.slLeafBottomA
    const slLeafBottomA = getProfileById(nodeProfileBinding.slLeafBottomA, 'SL оклейка A (низ)')
    const slLeafPerimeterBcd = getProfileById(
      nodeProfileBinding.slLeafPerimeterBcd,
      'SL оклейка B/C/D',
    )
    const slSideToFrame = getProfileById(nodeProfileBinding.slSideToFrame, 'SL рабочая сторона к раме')
    const slSideToFix = getProfileById(nodeProfileBinding.slSideToFix, 'SL рабочая сторона к FIX')
    const slSideToSl = getProfileById(nodeProfileBinding.slSideToSl, 'SL рабочая сторона к SL')
    const slSideToPenal = getProfileById(
      nodeProfileBinding.slSideToPenal,
      'SL рабочая сторона к пеналу',
    )
    const slSideToCorner = getProfileById(
      nodeProfileBinding.slSideToCorner,
      'SL рабочая сторона к углу',
    )

    const slCount = calculations.slidingPanels
    const slBottomTotal = slCount * panelWidth
    const slVerticalBdLengthPerLeaf = segmentHeightMm
    const slTopCLengthPerLeaf = panelWidth
    const slVerticalBdTotal = slCount * 2 * slVerticalBdLengthPerLeaf
    const slTopCTotal = slCount * slTopCLengthPerLeaf
    const slActiveVerticalLength = segmentHeightMm

    const slSideCounts = normalizedSegments.reduce(
      (acc, segment, index) => {
        if (segment !== 'SL') return acc
        const direction = normalizedOpeningDirections[index] ?? 'right'
        const adjacentIndex = direction === 'right' ? index + 1 : index - 1
        if (adjacentIndex < 0 || adjacentIndex >= normalizedSegments.length) {
          acc.frame += 1
          return acc
        }
        const adjacent = normalizedSegments[adjacentIndex]
        if (adjacent === 'FIX') {
          acc.fix += 1
          return acc
        }
        acc.sl += 1
        return acc
      },
      { frame: 0, fix: 0, sl: 0, penal: 0, corner: 0 },
    )

    const frameAssemblyBySystem: Record<FrameSystemType, string> = {
      'single-track': systemNodeTemplateBinding.frameSingleTrackAssembly,
      'double-track': systemNodeTemplateBinding.frameDoubleTrackAssembly,
      'triple-track': systemNodeTemplateBinding.frameTripleTrackAssembly,
      'quadruple-track': systemNodeTemplateBinding.frameQuadrupleTrackAssembly,
    }
    const selectedFrameAssemblyId =
      frameSystem.type === null ? '' : frameAssemblyBySystem[frameSystem.type]
    const frameTemplateItems: Omit<SpecItem, 'barsCount' | 'withWasteLengthMm'>[] = expandWithReadyAssembly(
      'frame-system-assembly',
      `Профиль рамы в сборе (${frameSystem.title})`,
      calculations.perimeter,
      selectedFrameAssemblyId,
    )

    const frameItems: Omit<SpecItem, 'barsCount' | 'withWasteLengthMm'>[] = frameComposition.map((item) => ({
      key: `frame-${item.part}-${item.profileName}`,
      article: item.article,
      profile: item.profileName,
      zone: `Состав рамы: ${item.part}`,
      partQuantity: 1,
      deductionMm: 0,
      additionMm: 0,
      totalLengthMm: item.totalLengthMm,
      stockLengthMm: item.stockLengthMm,
      sectionImageDataUrl: item.sectionImageDataUrl,
    }))

    const rawItems: Omit<SpecItem, 'barsCount' | 'withWasteLengthMm'>[] = [
      ...(frameTemplateItems.length > 0 ? frameTemplateItems : frameItems),
      ...customNodeSpecItems,
      ...expandWithNodeTemplate(
        'fix-fix',
        'Стыки FIX/FIX',
        calculations.fixFixJoints * segmentHeightMm,
        systemNodeTemplateBinding.fixFix,
        nodeProfileBinding.fixFix,
        fixFix.name,
      ),
      ...expandWithNodeTemplate(
        'fix-sl',
        'Стыки FIX/SL',
        calculations.fixSlJoints * segmentHeightMm,
        systemNodeTemplateBinding.fixSl,
        nodeProfileBinding.fixSl,
        fixSl.name,
      ),
      ...expandWithTemplateOrAssembly(
        'sl-leaf-bottom-a',
        leafBottomMode === 'freeway' ? 'SL оклейка A (низ, FreeWay)' : 'SL оклейка A (низ)',
        slBottomTotal,
        selectedSlBottomBindingId,
        nodeProfileBinding.slLeafBottomA,
        slLeafBottomA.name,
        true,
      ),
      ...expandWithTemplateOrAssembly(
        'sl-leaf-perimeter-bd',
        'SL оклейка B/D (вертикали)',
        slVerticalBdTotal,
        activeSlVerticalBdBindingId,
        nodeProfileBinding.slLeafPerimeterBcd,
        slLeafPerimeterBcd.name,
        true,
      ),
      ...expandWithTemplateOrAssembly(
        'sl-leaf-perimeter-c',
        'SL оклейка C (верх)',
        slTopCTotal,
        activeSlTopCBindingId,
        nodeProfileBinding.slLeafPerimeterBcd,
        slLeafPerimeterBcd.name,
        true,
      ),
      ...expandWithNodeTemplate(
        'sl-side-to-frame',
        'SL рабочая сторона: узел SL-рама',
        slSideCounts.frame * slActiveVerticalLength,
        systemNodeTemplateBinding.slSideToFrame,
        nodeProfileBinding.slSideToFrame,
        slSideToFrame.name,
      ),
      ...expandWithNodeTemplate(
        'sl-side-to-fix',
        'SL рабочая сторона: узел SL-FIX',
        slSideCounts.fix * slActiveVerticalLength,
        systemNodeTemplateBinding.slSideToFix,
        nodeProfileBinding.slSideToFix,
        slSideToFix.name,
      ),
      ...expandWithNodeTemplate(
        'sl-side-to-sl',
        'SL рабочая сторона: узел SL-SL',
        slSideCounts.sl * slActiveVerticalLength,
        systemNodeTemplateBinding.slSideToSl,
        nodeProfileBinding.slSideToSl,
        slSideToSl.name,
      ),
      ...expandWithNodeTemplate(
        'sl-side-to-penal',
        'SL рабочая сторона: узел SL-пенал',
        slSideCounts.penal * slActiveVerticalLength,
        systemNodeTemplateBinding.slSideToPenal,
        nodeProfileBinding.slSideToPenal,
        slSideToPenal.name,
      ),
      ...expandWithNodeTemplate(
        'sl-side-to-corner',
        'SL рабочая сторона: узел SL-угол',
        slSideCounts.corner * slActiveVerticalLength,
        systemNodeTemplateBinding.slSideToCorner,
        nodeProfileBinding.slSideToCorner,
        slSideToCorner.name,
      ),
    ]

    return rawItems
      .filter((item) => item.totalLengthMm > 0)
      .map((item) => {
        if (item.isHardware) {
          const withWasteLengthMm = item.totalLengthMm * (1 + wastePercent / 100)
          return {
            ...item,
            withWasteLengthMm,
            barsCount: Math.ceil(withWasteLengthMm - 1e-9),
          }
        }
        return {
          ...item,
          withWasteLengthMm: item.totalLengthMm * (1 + wastePercent / 100),
          barsCount: Math.ceil((item.totalLengthMm * (1 + wastePercent / 100)) / item.stockLengthMm),
        }
      })
  }, [
    calculations,
    customNodeSpecItems,
    frameSystem,
    frameComposition,
    hardware,
    readyHardware,
    nodeProfileBinding,
    panelWidth,
    profiles,
    normalizedOpeningDirections,
    normalizedSegments,
    leafBottomMode,
    nodeTemplates,
    readyAssemblies,
    segmentHeightMm,
    systemNodeTemplateBinding,
    wastePercent,
  ])

  const totalSpec = useMemo(() => {
    return spec.reduce(
      (acc, item) => {
        if (item.isHardware) {
          acc.hardwarePieces += item.totalLengthMm
          acc.hardwareWithWaste += item.withWasteLengthMm
          acc.hardwareBars += item.barsCount
        } else {
          acc.totalLengthMm += item.totalLengthMm
          acc.withWasteLengthMm += item.withWasteLengthMm
          acc.totalBars += item.barsCount
        }
        return acc
      },
      {
        totalLengthMm: 0,
        withWasteLengthMm: 0,
        totalBars: 0,
        hardwarePieces: 0,
        hardwareWithWaste: 0,
        hardwareBars: 0,
      },
    )
  }, [spec])

  const nodeProfileMap = useMemo(() => {
    const byId = (id: string) => profiles.find((p) => p.id === id)
    const templateById = (id: string) => nodeTemplates.find((p) => p.id === id)
    const assemblyById = (id: string) => readyAssemblies.find((p) => p.id === id)
    const displayNode = (templateId: string, fallbackProfileId: string) => {
      const template = templateById(templateId)
      if (template) return { article: template.article || 'УЗЕЛ', name: template.name }
      const assembly = assemblyById(templateId)
      if (assembly) return { article: assembly.article || 'СБОРКА', name: assembly.name }
      return byId(fallbackProfileId)
    }
    const rows = [
      {
        key: 'frameSingleTrackAssembly',
        node: 'Рама: однотрековый профиль',
        profile: assemblyById(systemNodeTemplateBinding.frameSingleTrackAssembly),
      },
      {
        key: 'frameDoubleTrackAssembly',
        node: 'Рама: двухтрековый профиль',
        profile: assemblyById(systemNodeTemplateBinding.frameDoubleTrackAssembly),
      },
      {
        key: 'frameTripleTrackAssembly',
        node: 'Рама: трехтрековый профиль',
        profile: assemblyById(systemNodeTemplateBinding.frameTripleTrackAssembly),
      },
      {
        key: 'frameQuadrupleTrackAssembly',
        node: 'Рама: четырехтрековый профиль',
        profile: assemblyById(systemNodeTemplateBinding.frameQuadrupleTrackAssembly),
      },
      {
        key: 'fixFix',
        node: 'Стык FIX/FIX',
        profile: displayNode(systemNodeTemplateBinding.fixFix, nodeProfileBinding.fixFix),
      },
      {
        key: 'fixSl',
        node: 'Стык FIX/SL',
        profile: displayNode(systemNodeTemplateBinding.fixSl, nodeProfileBinding.fixSl),
      },
      {
        key: 'slLeafBottomA',
        node: leafBottomMode === 'freeway' ? 'SL оклейка A (низ, FreeWay)' : 'SL оклейка A (низ)',
        profile: displayNode(
          leafBottomMode === 'freeway'
            ? systemNodeTemplateBinding.slLeafBottomAFreeway
            : systemNodeTemplateBinding.slLeafBottomA,
          nodeProfileBinding.slLeafBottomA,
        ),
      },
      {
        key: 'slLeafVerticalBd',
        node: 'SL оклейка B/D (вертикали)',
        profile: displayNode(
          systemNodeTemplateBinding.slLeafVerticalBd || systemNodeTemplateBinding.slLeafPerimeterBcd,
          nodeProfileBinding.slLeafPerimeterBcd,
        ),
      },
      {
        key: 'slLeafTopC',
        node: 'SL оклейка C (верх)',
        profile: displayNode(
          systemNodeTemplateBinding.slLeafTopC || systemNodeTemplateBinding.slLeafPerimeterBcd,
          nodeProfileBinding.slLeafPerimeterBcd,
        ),
      },
      {
        key: 'slSideToFrame',
        node: 'SL рабочая сторона -> рама',
        profile: displayNode(systemNodeTemplateBinding.slSideToFrame, nodeProfileBinding.slSideToFrame),
      },
      {
        key: 'slSideToFix',
        node: 'SL рабочая сторона -> FIX',
        profile: displayNode(systemNodeTemplateBinding.slSideToFix, nodeProfileBinding.slSideToFix),
      },
      {
        key: 'slSideToSl',
        node: 'SL рабочая сторона -> SL',
        profile: displayNode(systemNodeTemplateBinding.slSideToSl, nodeProfileBinding.slSideToSl),
      },
      {
        key: 'slSideToPenal',
        node: 'SL рабочая сторона -> пенал',
        profile: displayNode(systemNodeTemplateBinding.slSideToPenal, nodeProfileBinding.slSideToPenal),
      },
      {
        key: 'slSideToCorner',
        node: 'SL рабочая сторона -> угол',
        profile: displayNode(systemNodeTemplateBinding.slSideToCorner, nodeProfileBinding.slSideToCorner),
      },
    ]

    return rows.map((row) => ({
      node: row.node,
      article: row.profile?.article ?? '-',
      name: row.profile?.name ?? 'Профиль не найден в базе',
      status: row.profile ? 'OK' : 'Нет в базе',
    }))
  }, [nodeProfileBinding, profiles, nodeTemplates, readyAssemblies, systemNodeTemplateBinding, leafBottomMode])

  const buildProjectData = (): ProjectData => ({
    projectNumber,
    preparedBy,
    widthMm,
    heightMm,
    frameSeries,
    leafBottomMode,
    segmentsCount,
    segments: normalizedSegments,
    openingDirections: normalizedOpeningDirections,
    segmentHeightOffset65Mm,
    segmentHeightOffset80Mm,
    wastePercent,
    profiles,
    hardware,
    readyHardware,
    readyAssemblies,
    nodeTemplates,
    systemNodeTemplateBinding,
    nodeProfileBinding,
    customNodes,
  })

  const applyProjectData = (parsed: ProjectData) => {
    setProjectNumber(parsed.projectNumber ?? '')
    setPreparedBy(parsed.preparedBy ?? '')
    setWidthMm(parsed.widthMm)
    setHeightMm(parsed.heightMm)
    setFrameSeries(parsed.frameSeries ?? '65')
    setLeafBottomMode(parsed.leafBottomMode ?? 'standard')
    setSegmentsCount(Math.max(1, parsed.segmentsCount))
    setSegments(parsed.segments)
    setOpeningDirections(parsed.openingDirections ?? ['right', 'right', 'right'])
    setSegmentHeightOffset65Mm(Math.max(0, parsed.segmentHeightOffset65Mm ?? 52))
    setSegmentHeightOffset80Mm(Math.max(0, parsed.segmentHeightOffset80Mm ?? 67))
    setWastePercent(parsed.wastePercent ?? 0)
    setProfiles(parsed.profiles?.length ? parsed.profiles : defaultProfiles)
    setHardware(normalizeHardware(parsed.hardware))
    setReadyHardware(normalizeReadyHardware(parsed.readyHardware))
    setReadyAssemblies(normalizeReadyAssemblies(parsed.readyAssemblies))
    setNodeTemplates(normalizeNodeTemplates(parsed.nodeTemplates))
    setSystemNodeTemplateBinding({
      ...defaultSystemNodeTemplateBinding,
      ...(parsed.systemNodeTemplateBinding ?? {}),
    })
    setNodeProfileBinding({ ...defaultNodeProfileBinding, ...(parsed.nodeProfileBinding ?? {}) })
    setCustomNodes(normalizeCustomNodes(parsed.customNodes))
  }

  const saveProjectToServer = async () => {
    setServerSyncStatus('Сохраняю на сервер...')
    try {
      const response = await fetch('/api/project', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(buildProjectData()),
      })
      if (!response.ok) throw new Error('API error')
      setServerSyncStatus('Сохранено на сервер')
    } catch {
      setServerSyncStatus('Ошибка сохранения на сервер')
    }
  }

  useEffect(() => {
    const bootstrap = async () => {
      setServerSyncStatus('Загружаю базу с сервера...')
      try {
        const response = await fetch('/api/project')
        if (!response.ok) throw new Error('API error')
        const payload = (await response.json()) as { data?: ProjectData | null }
        if (payload.data) {
          applyProjectData(payload.data)
          setServerSyncStatus('База загружена с сервера')
        } else {
          setServerSyncStatus('На сервере пока нет данных, используется шаблон')
        }
      } catch {
        setServerSyncStatus('Ошибка загрузки с сервера, используется шаблон')
      } finally {
        setServerLoadReady(true)
      }
    }

    void bootstrap()
  }, [])

  useEffect(() => {
    if (!serverLoadReady) return
    if (autosaveTimerRef.current) window.clearTimeout(autosaveTimerRef.current)
    autosaveTimerRef.current = window.setTimeout(() => {
      void saveProjectToServer()
    }, 600)

    return () => {
      if (autosaveTimerRef.current) {
        window.clearTimeout(autosaveTimerRef.current)
        autosaveTimerRef.current = undefined
      }
    }
  }, [
    serverLoadReady,
    heightMm,
    normalizedOpeningDirections,
    normalizedSegments,
    nodeProfileBinding,
    systemNodeTemplateBinding,
    customNodes,
    profiles,
    readyHardware,
    readyAssemblies,
    nodeTemplates,
    segmentsCount,
    segmentHeightOffset65Mm,
    segmentHeightOffset80Mm,
    wastePercent,
    widthMm,
    frameSeries,
    leafBottomMode,
    projectNumber,
    preparedBy,
    hardware,
    readyHardware,
    activeTab,
    newProfileArticle,
    newProfileName,
    newProfileLength,
  ])

  useEffect(() => {
    if (typeof window === 'undefined') return
    const nextHash = `#${activeTab}`
    if (window.location.hash !== nextHash) {
      window.history.replaceState(null, '', nextHash)
    }
  }, [activeTab])

  const exportToExcel = () => {
    const getFormulaText = (item: SpecItem) => {
      if (item.isHardware) {
        return `Фурнитура, итого шт: ${item.totalLengthMm.toFixed(0)}`
      }
      const roundedHeight = Number(segmentHeightMm.toFixed(0))
      const roundedPanel = Number(panelWidth.toFixed(0))
      if (item.zone.startsWith('SL оклейка A (низ)')) {
        return `${calculations.slidingPanels} × ${item.partQuantity} × (${roundedPanel} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
      }
      if (item.zone.startsWith('SL оклейка B/D (вертикали)')) {
        return `${calculations.slidingPanels} × 2 × ${item.partQuantity} × (${roundedHeight} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
      }
      if (item.zone.startsWith('SL оклейка C (верх)')) {
        return `${calculations.slidingPanels} × ${item.partQuantity} × (${roundedPanel} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
      }
      if (item.zone.startsWith('Стыки FIX/FIX')) {
        return `${calculations.fixFixJoints} × ${item.partQuantity} × (${roundedHeight} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
      }
      if (item.zone.startsWith('Стыки FIX/SL')) {
        return `${calculations.fixSlJoints} × ${item.partQuantity} × (${roundedHeight} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
      }
      return `1 × ${item.totalLengthMm.toFixed(0)} = ${item.totalLengthMm.toFixed(0)}`
    }

    const rows = spec.map((item) => ({
      '№ проекта': projectNumber || '-',
      Составил: preparedBy || '-',
      Артикул: item.article,
      Профиль: item.profile,
      Назначение: item.zone,
      'Формула расчета': getFormulaText(item),
      'Кол-во в сборке, шт': item.partQuantity,
      'Суммарная длина, мм': Number(item.totalLengthMm.toFixed(0)),
      'С учетом отхода, мм': Number(item.withWasteLengthMm.toFixed(0)),
      'Длина хлыста, мм': item.isHardware ? '—' : item.stockLengthMm,
      'Необходимо хлыстов, шт': item.barsCount,
    }))

    const worksheet = XLSX.utils.json_to_sheet(rows)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Спецификация')
    XLSX.writeFile(workbook, 'specification.xlsx')
  }

  const exportToPdf = () => {
    const getFormulaText = (item: SpecItem) => {
      if (item.isHardware) {
        return `Фурнитура, итого шт: ${item.totalLengthMm.toFixed(0)}`
      }
      const roundedHeight = Number(segmentHeightMm.toFixed(0))
      const roundedPanel = Number(panelWidth.toFixed(0))
      if (item.zone.startsWith('SL оклейка A (низ)')) {
        return `${calculations.slidingPanels} × ${item.partQuantity} × (${roundedPanel} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
      }
      if (item.zone.startsWith('SL оклейка B/D (вертикали)')) {
        return `${calculations.slidingPanels} × 2 × ${item.partQuantity} × (${roundedHeight} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
      }
      if (item.zone.startsWith('SL оклейка C (верх)')) {
        return `${calculations.slidingPanels} × ${item.partQuantity} × (${roundedPanel} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
      }
      if (item.zone.startsWith('Стыки FIX/FIX')) {
        return `${calculations.fixFixJoints} × ${item.partQuantity} × (${roundedHeight} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
      }
      if (item.zone.startsWith('Стыки FIX/SL')) {
        return `${calculations.fixSlJoints} × ${item.partQuantity} × (${roundedHeight} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
      }
      return `1 × ${item.totalLengthMm.toFixed(0)} = ${item.totalLengthMm.toFixed(0)}`
    }

    const doc = new jsPDF()
    doc.setFontSize(14)
    doc.text('Спецификация профилей', 14, 16)
    doc.setFontSize(10)
    doc.text(`Рама: ${widthMm} x ${heightMm} мм`, 14, 22)
    doc.text(`Высота сегмента: ${segmentHeightMm.toFixed(0)} мм`, 14, 27)
    doc.text(`Отход: ${wastePercent}%`, 14, 32)
    doc.text(`Сегменты: ${normalizedSegments.join(' | ')}`, 14, 37)

    doc.text(`№ проекта: ${projectNumber || '-'}`, 14, 42)
    doc.text(`Составил: ${preparedBy || '-'}`, 14, 47)

    autoTable(doc, {
      startY: 53,
      head: [[
        'Артикул',
        'Профиль',
        'Назначение',
        'Формула',
        'Кол-во в сборке, шт',
        'Суммарная длина, мм',
        'С учетом отхода, мм',
        'Длина хлыста, мм',
        'Хлысты, шт',
      ]],
      body: spec.map((item) => [
        item.article,
        item.profile,
        item.zone,
        getFormulaText(item),
        String(item.partQuantity),
        item.totalLengthMm.toFixed(0),
        item.withWasteLengthMm.toFixed(0),
        item.isHardware ? '—' : String(item.stockLengthMm),
        String(item.barsCount),
      ]),
      styles: {
        fontSize: 9,
      },
    })

    doc.save('specification.pdf')
  }

  const saveProject = () => {
    const data = buildProjectData()

    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = 'frame-project.json'
    link.click()
    URL.revokeObjectURL(url)
  }

  const loadProject = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = () => {
      try {
        const parsed = JSON.parse(String(reader.result)) as ProjectData
        applyProjectData(parsed)
      } catch {
        alert('Не удалось загрузить JSON проекта')
      }
    }
    reader.readAsText(file)
    event.target.value = ''
  }

  const setSegmentType = (index: number, value: SegmentType) => {
    if (frameSeries === '65' && value === 'SL') {
      showWarning(
        'Серия 65 не поддерживает откатные створки. Доступны только FIX сегменты.',
        'series65-block-sl',
      )
      return
    }
    setSegments((current) => {
      const next = [...current]
      next[index] = value
      return next
    })
  }

  const setOpeningDirection = (index: number, value: OpeningDirection) => {
    setOpeningDirections((current) => {
      const next = [...current]
      next[index] = value
      return next
    })
  }

  const handleNewProfileSectionFile = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      setNewProfileSectionImageDataUrl(String(reader.result))
    }
    reader.readAsDataURL(file)
    event.target.value = ''
  }

  const handleProfileSectionFile = (id: string, event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      const url = String(reader.result)
      updateProfile(id, { sectionImageDataUrl: url })
    }
    reader.readAsDataURL(file)
    event.target.value = ''
  }

  const handleNewAssemblySectionFile = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      setNewAssemblySectionImageDataUrl(String(reader.result))
    }
    reader.readAsDataURL(file)
    event.target.value = ''
  }

  const handleReadyAssemblySectionFile = (id: string, event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      const url = String(reader.result)
      updateReadyAssembly(id, { sectionImageDataUrl: url })
    }
    reader.readAsDataURL(file)
    event.target.value = ''
  }

  const addProfile = () => {
    if (!newProfileName.trim()) return
    const nextId = `profile-${Date.now()}`
    setProfiles((current) => [
      ...current,
      {
        id: nextId,
        article: newProfileArticle.trim(),
        name: newProfileName.trim(),
        defaultLength: newProfileLength,
        sectionImageDataUrl: newProfileSectionImageDataUrl || undefined,
      },
    ])
    setNewProfileName('')
    setNewProfileArticle('')
    setNewProfileLength(3000)
    setNewProfileSectionImageDataUrl('')
  }

  const updateProfile = (id: string, patch: Partial<Profile>) => {
    setProfiles((current) =>
      current.map((profile) => (profile.id === id ? { ...profile, ...patch } : profile)),
    )
  }

  const deleteProfile = (id: string) => {
    setProfiles((current) => {
      if (current.length <= 1) return current
      const next = current.filter((profile) => profile.id !== id)
      setCustomNodes((prev) =>
        prev
          .map((node) => ({
            ...node,
            parts: node.parts.filter((part) => part.profileId !== id),
          }))
          .filter((node) => node.parts.length > 0),
      )
      setReadyAssemblies((prev) =>
        prev
          .map((assembly) => ({
            ...assembly,
            parts: assembly.parts.filter((part) => part.profileId !== id),
          }))
          .filter((assembly) => assembly.parts.length > 0),
      )
      setNodeTemplates((prev) =>
        prev
          .map((template) => ({
            ...template,
            parts: template.parts.filter((part) => part.profileId !== id),
          }))
          .filter((template) => template.parts.length > 0),
      )
      setNodeProfileBinding((prev) => ({
        frameOuter65: prev.frameOuter65 === id ? '' : prev.frameOuter65,
        frameOuter80: prev.frameOuter80 === id ? '' : prev.frameOuter80,
        frameThermal: prev.frameThermal === id ? '' : prev.frameThermal,
        frameCentral65: prev.frameCentral65 === id ? '' : prev.frameCentral65,
        frameCentral80: prev.frameCentral80 === id ? '' : prev.frameCentral80,
        fixFix: prev.fixFix === id ? '' : prev.fixFix,
        fixSl: prev.fixSl === id ? '' : prev.fixSl,
        slLeafBottomA: prev.slLeafBottomA === id ? '' : prev.slLeafBottomA,
        slLeafPerimeterBcd: prev.slLeafPerimeterBcd === id ? '' : prev.slLeafPerimeterBcd,
        slSideToFrame: prev.slSideToFrame === id ? '' : prev.slSideToFrame,
        slSideToFix: prev.slSideToFix === id ? '' : prev.slSideToFix,
        slSideToSl: prev.slSideToSl === id ? '' : prev.slSideToSl,
        slSideToPenal: prev.slSideToPenal === id ? '' : prev.slSideToPenal,
        slSideToCorner: prev.slSideToCorner === id ? '' : prev.slSideToCorner,
      }))
      return next
    })
  }

  const addHardware = () => {
    if (!newHardwareArticle.trim() && !newHardwareName.trim()) return
    setHardware((current) => [
      ...current,
      {
        id: `hardware-${Date.now()}`,
        article: newHardwareArticle.trim(),
        name: newHardwareName.trim(),
      },
    ])
    setNewHardwareArticle('')
    setNewHardwareName('')
  }

  const updateHardware = (id: string, patch: Partial<HardwareItem>) => {
    setHardware((current) =>
      current.map((item) => (item.id === id ? { ...item, ...patch } : item)),
    )
  }

  const duplicateHardware = (id: string) => {
    setHardware((current) => {
      const source = current.find((item) => item.id === id)
      if (!source) return current
      return [
        ...current,
        {
          ...source,
          id: `hardware-${Date.now()}`,
          name: `${source.name} (копия)`,
        },
      ]
    })
  }

  const deleteHardware = (id: string) => {
    setHardware((current) => current.filter((item) => item.id !== id))
    setCustomNodes((prev) =>
      prev
        .map((node) => ({
          ...node,
          parts: node.parts.filter((part) => part.hardwareId !== id),
        }))
        .filter((node) => node.parts.length > 0),
    )
    setReadyAssemblies((prev) =>
      prev
        .map((assembly) => ({
          ...assembly,
          parts: assembly.parts.filter((part) => part.hardwareId !== id),
        }))
        .filter((assembly) => assembly.parts.length > 0),
    )
    setNodeTemplates((prev) =>
      prev
        .map((template) => ({
          ...template,
          parts: template.parts.filter((part) => part.hardwareId !== id),
        }))
        .filter((template) => template.parts.length > 0),
    )
  }

  const addReadyHardware = () => {
    if (!newReadyHardwareArticle.trim() && !newReadyHardwareName.trim()) return
    setReadyHardware((current) => [
      ...current,
      {
        id: `ready-hardware-${Date.now()}`,
        article: newReadyHardwareArticle.trim(),
        name: newReadyHardwareName.trim(),
      },
    ])
    setNewReadyHardwareArticle('')
    setNewReadyHardwareName('')
  }

  const updateReadyHardware = (id: string, patch: Partial<ReadyHardwareItem>) => {
    setReadyHardware((current) =>
      current.map((item) => (item.id === id ? { ...item, ...patch } : item)),
    )
  }

  const duplicateReadyHardware = (id: string) => {
    setReadyHardware((current) => {
      const source = current.find((item) => item.id === id)
      if (!source) return current
      return [
        ...current,
        {
          ...source,
          id: `ready-hardware-${Date.now()}`,
          name: `${source.name} (копия)`,
        },
      ]
    })
  }

  const deleteReadyHardware = (id: string) => {
    setReadyHardware((current) => current.filter((item) => item.id !== id))
    setCustomNodes((prev) =>
      prev
        .map((node) => ({
          ...node,
          parts: node.parts.filter((part) => part.readyHardwareId !== id),
        }))
        .filter((node) => node.parts.length > 0),
    )
    setReadyAssemblies((prev) =>
      prev
        .map((assembly) => ({
          ...assembly,
          parts: assembly.parts.filter((part) => part.readyHardwareId !== id),
        }))
        .filter((assembly) => assembly.parts.length > 0),
    )
    setNodeTemplates((prev) =>
      prev
        .map((template) => ({
          ...template,
          parts: template.parts.filter((part) => part.readyHardwareId !== id),
        }))
        .filter((template) => template.parts.length > 0),
    )
  }

  const duplicateProfile = (id: string) => {
    setProfiles((current) => {
      const source = current.find((profile) => profile.id === id)
      if (!source) return current
      return [
        ...current,
        {
          ...source,
          id: `profile-${Date.now()}`,
          name: `${source.name} (копия)`,
        },
      ]
    })
  }

  const addCustomNode = () => {
    const preparedParts = newNodeParts
      .filter((part) => part.profileId || part.hardwareId || part.readyHardwareId)
      .map((part) => ({
        ...part,
        quantity: Math.max(1, Number(part.quantity) || 1),
        deductionMm: Math.max(0, part.deductionMm),
        additionMm: Math.max(0, part.additionMm),
      }))
    if (!newNodeName.trim() || preparedParts.length === 0) return
    setCustomNodes((current) => [
      ...current,
      {
        id: `${Date.now()}`,
        name: newNodeName.trim(),
        formula: newNodeFormula,
        quantityMode: newNodeQuantityMode,
        quantityValue: Math.max(1, newNodeQuantityValue),
        parts: preparedParts,
      },
    ])
    setNewNodeName('')
    setNewNodeQuantityValue(1)
    setNewNodeParts([
      {
        id: `new-part-${Date.now()}`,
        profileId: '',
        hardwareId: '',
        readyHardwareId: '',
        quantity: 1,
        deductionMm: 0,
        additionMm: 0,
      },
    ])
  }

  const removeCustomNode = (id: string) => {
    setCustomNodes((current) => current.filter((node) => node.id !== id))
  }

  const duplicateCustomNode = (id: string) => {
    setCustomNodes((current) => {
      const source = current.find((node) => node.id === id)
      if (!source) return current
      const duplicatedId = `node-${Date.now()}`
      return [
        ...current,
        {
          ...source,
          id: duplicatedId,
          name: `${source.name} (копия)`,
          parts: source.parts.map((part, index) => ({
            ...part,
            id: `${duplicatedId}-part-${index}`,
          })),
        },
      ]
    })
  }

  const updateCustomNode = (id: string, patch: Partial<CustomNode>) => {
    setCustomNodes((current) =>
      current.map((node) => (node.id === id ? { ...node, ...patch } : node)),
    )
  }

  const addPartToCustomNode = (nodeId: string) => {
    setCustomNodes((current) =>
      current.map((node) =>
        node.id === nodeId
          ? {
              ...node,
              parts: [
                ...node.parts,
                {
                  id: `part-${Date.now()}-${node.parts.length}`,
                  profileId: '',
                  hardwareId: '',
                  readyHardwareId: '',
                  quantity: 1,
                  deductionMm: 0,
                  additionMm: 0,
                },
              ],
            }
          : node,
      ),
    )
  }

  const updatePartInCustomNode = (nodeId: string, partId: string, patch: Partial<CustomNodePart>) => {
    setCustomNodes((current) =>
      current.map((node) =>
        node.id === nodeId
          ? {
              ...node,
              parts: node.parts.map((part) => (part.id === partId ? { ...part, ...patch } : part)),
            }
          : node,
      ),
    )
  }

  const removePartFromCustomNode = (nodeId: string, partId: string) => {
    setCustomNodes((current) =>
      current.map((node) =>
        node.id === nodeId
          ? {
              ...node,
              parts:
                node.parts.length <= 1
                  ? node.parts
                  : node.parts.filter((part) => part.id !== partId),
            }
          : node,
      ),
    )
  }

  const addNewNodePart = () => {
    setNewNodeParts((current) => [
      ...current,
      {
        id: `new-part-${Date.now()}-${current.length}`,
        profileId: '',
        hardwareId: '',
        readyHardwareId: '',
        quantity: 1,
        deductionMm: 0,
        additionMm: 0,
      },
    ])
  }

  const updateNewNodePart = (partId: string, patch: Partial<CustomNodePart>) => {
    setNewNodeParts((current) =>
      current.map((part) => (part.id === partId ? { ...part, ...patch } : part)),
    )
  }

  const removeNewNodePart = (partId: string) => {
    setNewNodeParts((current) => current.filter((part) => part.id !== partId))
  }

  const addReadyAssembly = () => {
    if (!newAssemblyName.trim()) {
      showWarning('Введите название готового профиля перед добавлением.', 'add-assembly-missing-name')
      return
    }
    const normalizedParts = newAssemblyParts.map((part) =>
      resolveAssemblyPartSelection(part, newAssemblyPartSearch[part.id] ?? ''),
    )
    const parts = normalizedParts.filter(
      (part) => part.profileId || part.hardwareId || part.readyHardwareId,
    )
    if (parts.length === 0) {
      showWarning(
        'Добавьте хотя бы одну деталь и выберите позицию из выпадающего списка (профиль/фурнитура/готовая фурнитура).',
        'add-assembly-missing-parts',
      )
      return
    }
    setReadyAssemblies((current) => [
      ...current,
      {
        id: `assembly-${Date.now()}`,
        article: newAssemblyArticle.trim(),
        name: newAssemblyName.trim(),
        sectionImageDataUrl: newAssemblySectionImageDataUrl || undefined,
        parts: parts.map((part) => ({
          ...part,
          quantity: Math.max(1, Number(part.quantity) || 1),
          deductionMm: Math.max(0, part.deductionMm),
          additionMm: Math.max(0, part.additionMm),
        })),
      },
    ])
    setNewAssemblyArticle('')
    setNewAssemblyName('')
    setNewAssemblySectionImageDataUrl('')
    setNewAssemblyParts([
      {
        id: `new-assembly-part-${Date.now()}`,
        profileId: '',
        hardwareId: '',
        readyHardwareId: '',
        quantity: 1,
        deductionMm: 0,
        additionMm: 0,
      },
    ])
    setNewAssemblyPartSearch({})
  }

  const updateReadyAssembly = (id: string, patch: Partial<ReadyAssembly>) => {
    setReadyAssemblies((current) =>
      current.map((assembly) => (assembly.id === id ? { ...assembly, ...patch } : assembly)),
    )
  }

  const removeReadyAssembly = (id: string) => {
    setReadyAssemblies((current) => current.filter((assembly) => assembly.id !== id))
    setReadyAssemblyPartSearch((current) => {
      const next = { ...current }
      Object.keys(next).forEach((key) => {
        if (key.startsWith(`${id}::`)) delete next[key]
      })
      return next
    })
    setSystemNodeTemplateBinding((current) => ({
      ...current,
      frameSingleTrackAssembly:
        current.frameSingleTrackAssembly === id ? '' : current.frameSingleTrackAssembly,
      frameDoubleTrackAssembly:
        current.frameDoubleTrackAssembly === id ? '' : current.frameDoubleTrackAssembly,
      frameTripleTrackAssembly:
        current.frameTripleTrackAssembly === id ? '' : current.frameTripleTrackAssembly,
      frameQuadrupleTrackAssembly:
        current.frameQuadrupleTrackAssembly === id ? '' : current.frameQuadrupleTrackAssembly,
      slLeafBottomA: current.slLeafBottomA === id ? '' : current.slLeafBottomA,
      slLeafBottomAFreeway:
        current.slLeafBottomAFreeway === id ? '' : current.slLeafBottomAFreeway,
      slLeafVerticalBd: current.slLeafVerticalBd === id ? '' : current.slLeafVerticalBd,
      slLeafTopC: current.slLeafTopC === id ? '' : current.slLeafTopC,
    }))
  }

  const addPartToReadyAssembly = (assemblyId: string) => {
    let newPartId = ''
    setReadyAssemblies((current) =>
      current.map((assembly) => {
        if (assembly.id !== assemblyId) return assembly
        newPartId = `${assemblyId}-part-${Date.now()}-${assembly.parts.length}`
        return {
          ...assembly,
          parts: [
            ...assembly.parts,
            {
              id: newPartId,
              profileId: '',
              hardwareId: '',
              readyHardwareId: '',
              quantity: 1,
              deductionMm: 0,
              additionMm: 0,
            },
          ],
        }
      }),
    )
    if (newPartId) {
      const searchKey = `${assemblyId}::${newPartId}`
      setReadyAssemblyPartSearch((current) => ({ ...current, [searchKey]: '' }))
    }
  }

  const updatePartInReadyAssembly = (
    assemblyId: string,
    partId: string,
    patch: Partial<CustomNodePart>,
  ) => {
    setReadyAssemblies((current) =>
      current.map((assembly) =>
        assembly.id === assemblyId
          ? {
              ...assembly,
              parts: assembly.parts.map((part) =>
                part.id === partId ? { ...part, ...patch } : part,
              ),
            }
          : assembly,
      ),
    )
  }

  const removePartFromReadyAssembly = (assemblyId: string, partId: string) => {
    const searchKey = `${assemblyId}::${partId}`
    setReadyAssemblies((current) =>
      current.map((assembly) =>
        assembly.id === assemblyId
          ? {
              ...assembly,
              parts:
                assembly.parts.length <= 1
                  ? assembly.parts
                  : assembly.parts.filter((part) => part.id !== partId),
            }
          : assembly,
      ),
    )
    setReadyAssemblyPartSearch((current) => {
      if (!(searchKey in current)) return current
      const next = { ...current }
      delete next[searchKey]
      return next
    })
  }

  const addNewAssemblyPart = () => {
    let newPartId = ''
    setNewAssemblyParts((current) => {
      newPartId = `new-assembly-part-${Date.now()}-${current.length}`
      return [
        ...current,
        {
          id: newPartId,
          profileId: '',
          hardwareId: '',
          readyHardwareId: '',
          quantity: 1,
          deductionMm: 0,
          additionMm: 0,
        },
      ]
    })
    if (newPartId) {
      setNewAssemblyPartSearch((current) => ({ ...current, [newPartId]: '' }))
    }
  }

  const updateNewAssemblyPart = (partId: string, patch: Partial<CustomNodePart>) => {
    setNewAssemblyParts((current) =>
      current.map((part) => (part.id === partId ? { ...part, ...patch } : part)),
    )
  }

  const removeNewAssemblyPart = (partId: string) => {
    setNewAssemblyParts((current) =>
      current.length <= 1 ? current : current.filter((part) => part.id !== partId),
    )
    setNewAssemblyPartSearch((current) => {
      if (!(partId in current)) return current
      const next = { ...current }
      delete next[partId]
      return next
    })
  }

  const addNodeTemplate = () => {
    const parts = newNodeTemplateParts.filter(
      (part) => part.profileId || part.hardwareId || part.readyHardwareId,
    )
    if (!newNodeTemplateName.trim() || parts.length === 0) return
    setNodeTemplates((current) => [
      ...current,
      {
        id: `node-template-${Date.now()}`,
        article: newNodeTemplateArticle.trim(),
        name: newNodeTemplateName.trim(),
        parts: parts.map((part) => ({
          ...part,
          quantity: Math.max(1, Number(part.quantity) || 1),
          deductionMm: Math.max(0, part.deductionMm),
          additionMm: Math.max(0, part.additionMm),
        })),
      },
    ])
    setNewNodeTemplateArticle('')
    setNewNodeTemplateName('')
    setNewNodeTemplateParts([
      {
        id: `new-node-template-part-${Date.now()}`,
        profileId: '',
        hardwareId: '',
        readyHardwareId: '',
        quantity: 1,
        deductionMm: 0,
        additionMm: 0,
      },
    ])
  }

  const updateNodeTemplate = (id: string, patch: Partial<NodeTemplate>) => {
    setNodeTemplates((current) =>
      current.map((template) => (template.id === id ? { ...template, ...patch } : template)),
    )
  }

  const removeNodeTemplate = (id: string) => {
    setNodeTemplates((current) => current.filter((template) => template.id !== id))
  }

  const addNewNodeTemplatePart = () => {
    setNewNodeTemplateParts((current) => [
      ...current,
      {
        id: `new-node-template-part-${Date.now()}-${current.length}`,
        profileId: '',
        hardwareId: '',
        readyHardwareId: '',
        quantity: 1,
        deductionMm: 0,
        additionMm: 0,
      },
    ])
  }

  const updateNewNodeTemplatePart = (partId: string, patch: Partial<CustomNodePart>) => {
    setNewNodeTemplateParts((current) =>
      current.map((part) => (part.id === partId ? { ...part, ...patch } : part)),
    )
  }

  const removeNewNodeTemplatePart = (partId: string) => {
    setNewNodeTemplateParts((current) =>
      current.length <= 1 ? current : current.filter((part) => part.id !== partId),
    )
  }

  const addPartToNodeTemplate = (templateId: string) => {
    setNodeTemplates((current) =>
      current.map((template) =>
        template.id === templateId
          ? {
              ...template,
              parts: [
                ...template.parts,
                {
                  id: `${templateId}-part-${Date.now()}-${template.parts.length}`,
                  profileId: '',
                  hardwareId: '',
                  readyHardwareId: '',
                  quantity: 1,
                  deductionMm: 0,
                  additionMm: 0,
                },
              ],
            }
          : template,
      ),
    )
  }

  const updatePartInNodeTemplate = (
    templateId: string,
    partId: string,
    patch: Partial<CustomNodePart>,
  ) => {
    setNodeTemplates((current) =>
      current.map((template) =>
        template.id === templateId
          ? {
              ...template,
              parts: template.parts.map((part) => (part.id === partId ? { ...part, ...patch } : part)),
            }
          : template,
      ),
    )
  }

  const removePartFromNodeTemplate = (templateId: string, partId: string) => {
    setNodeTemplates((current) =>
      current.map((template) =>
        template.id === templateId
          ? {
              ...template,
              parts:
                template.parts.length <= 1
                  ? template.parts
                  : template.parts.filter((part) => part.id !== partId),
            }
          : template,
      ),
    )
  }

  const formulaLabels: Record<CustomNodeFormula, string> = {
    framePerimeter: 'Периметр рамы',
    frameTop: 'Верх рамы',
    frameSidesBottom: 'Бока + низ рамы',
    frameHeight: 'Высота рамы',
    panelWidth: 'Ширина сегмента',
    slPerimeterSingle: 'Периметр одной SL створки',
  }

  const quantityModeLabels: Record<CustomNodeQuantityMode, string> = {
    fixed: 'Фиксированное',
    slidingCount: 'Количество SL',
    segmentCount: 'Количество сегментов',
    verticalJoints: 'Количество вертикальных стыков',
  }

  const addDefaultSlNodeTemplates = () => {
    const existingNames = new Set(nodeTemplates.map((item) => item.name.trim().toLowerCase()))
    const seedProfileId = profiles[0]?.id ?? ''
    const seedHardwareId = hardware[0]?.id ?? ''
    const created: NodeTemplate[] = []

    defaultSlNodeTemplateNames.forEach((name) => {
      if (existingNames.has(name.toLowerCase())) return
      created.push({
        id: `node-template-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        article: '',
        name,
        parts: [
          {
            id: `part-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
            profileId: seedProfileId,
            hardwareId: seedProfileId ? '' : seedHardwareId,
            readyHardwareId: '',
            quantity: 1,
            deductionMm: 0,
            additionMm: 0,
          },
        ],
      })
    })

    if (created.length === 0) {
      showWarning('Стандартные SL-узлы уже добавлены в базу.', 'sl-node-templates-exist')
      return
    }

    setNodeTemplates((current) => [...current, ...created])
    showWarning(
      `Добавлено узлов: ${created.map((item) => item.name).join(', ')}. При необходимости отредактируйте состав.`,
      'sl-node-templates-created',
    )
  }

  const resetProjectToTemplate = () => {
    if (!window.confirm('Сбросить проект к шаблону по умолчанию?')) return
    setActiveTab('config')
    setProjectNumber('')
    setPreparedBy('')
    setWidthMm(6000)
    setHeightMm(3000)
    setFrameSeries('65')
    setLeafBottomMode('standard')
    setUiWarning('')
    setSegmentsCount(3)
    setSegments(['FIX', 'FIX', 'SL'])
    setOpeningDirections(['right', 'right', 'right'])
    setWastePercent(7)
    setProfiles(defaultProfiles)
    setHardware([])
    setNewHardwareArticle('')
    setNewHardwareName('')
    setHardwareSearch('')
    setReadyHardware([])
    setNewReadyHardwareArticle('')
    setNewReadyHardwareName('')
    setReadyHardwareSearch('')
    setReadyAssemblies([])
    setNodeTemplates([])
    setSystemNodeTemplateBinding(defaultSystemNodeTemplateBinding)
    setNodeProfileBinding(defaultNodeProfileBinding)
    setCustomNodes([])
    setNewProfileName('')
    setNewProfileArticle('')
    setNewProfileLength(3000)
    setNewProfileSectionImageDataUrl('')
    setNewNodeName('')
    setNewNodeFormula('framePerimeter')
    setNewNodeQuantityMode('fixed')
    setNewNodeQuantityValue(1)
    setNewNodeParts([
      {
        id: 'new-part-1',
        profileId: '',
        hardwareId: '',
        readyHardwareId: '',
        quantity: 1,
        deductionMm: 0,
        additionMm: 0,
      },
    ])
    setNewAssemblyArticle('')
    setNewAssemblyName('')
    setNewAssemblySectionImageDataUrl('')
    setNewAssemblyParts([
      {
        id: 'new-assembly-part-1',
        profileId: '',
        hardwareId: '',
        readyHardwareId: '',
        quantity: 1,
        deductionMm: 0,
        additionMm: 0,
      },
    ])
    setNewNodeTemplateArticle('')
    setNewNodeTemplateName('')
    setNewNodeTemplateParts([
      {
        id: 'new-node-template-part-1',
        profileId: '',
        hardwareId: '',
        readyHardwareId: '',
        quantity: 1,
        deductionMm: 0,
        additionMm: 0,
      },
    ])
    setServerSyncStatus('Шаблон применён, сохраняю на сервер...')
  }

  return (
    <main className="app">
      <header>
        <h1>Конфигуратор рам и стеклопакетов</h1>
        <p>Локальный MVP: размеры, сегменты, схема и первичный расчет профилей</p>
      </header>

      <div className="tabs">
        <button
          type="button"
          className={activeTab === 'config' ? 'active' : ''}
          onClick={() => setActiveTab('config')}
        >
          1. Конфигуратор
        </button>
        <button
          type="button"
          className={activeTab === 'profiles' ? 'active' : ''}
          onClick={() => setActiveTab('profiles')}
        >
          2. База профилей
        </button>
        <button
          type="button"
          className={activeTab === 'hardware' ? 'active' : ''}
          onClick={() => setActiveTab('hardware')}
        >
          3. База фурнитуры
        </button>
        <button
          type="button"
          className={activeTab === 'ready-hardware' ? 'active' : ''}
          onClick={() => setActiveTab('ready-hardware')}
        >
          4. Готовая фурнитура
        </button>
        <button
          type="button"
          className={activeTab === 'assemblies' ? 'active' : ''}
          onClick={() => setActiveTab('assemblies')}
        >
          5. Готовые профили
        </button>
        <button
          type="button"
          className={activeTab === 'nodes' ? 'active' : ''}
          onClick={() => setActiveTab('nodes')}
        >
          6. База узлов
        </button>
        <button
          type="button"
          className={activeTab === 'settings' ? 'active' : ''}
          onClick={() => setActiveTab('settings')}
        >
          7. Настройки
        </button>
      </div>
      {uiWarning ? <p className="warning-text">{uiWarning}</p> : null}

      {activeTab === 'config' ? (
        <section className="content">
          <div className="grid">
            <label>
              № проекта
              <input
                type="text"
                value={projectNumber}
                onChange={(e) => setProjectNumber(e.target.value)}
                placeholder="Например: VP-2026-041"
              />
            </label>
            <label>
              Составил
              <input
                type="text"
                value={preparedBy}
                onChange={(e) => setPreparedBy(e.target.value)}
                placeholder="ФИО"
              />
            </label>
          </div>
          <div className="grid">
            <label>
              Ширина рамы, мм
              <input
                type="number"
                value={widthMm}
                min={100}
                onChange={(e) => setWidthMm(Number(e.target.value) || 0)}
              />
            </label>
            <label>
              Высота рамы, мм
              <input
                type="number"
                value={heightMm}
                min={100}
                onChange={(e) => setHeightMm(Number(e.target.value) || 0)}
              />
            </label>
            <label>
              Серия верхнего профиля
              <select
                value={frameSeries}
                onChange={(e) => setFrameSeries(e.target.value as FrameSeries)}
              >
                <option value="65">65</option>
                <option value="80">80</option>
              </select>
              <small className="hint">Боковые и нижний профили рамы фиксировано 65.</small>
            </label>
            <label>
              Тип нижнего профиля створки
              <select
                value={leafBottomMode}
                onChange={(e) => setLeafBottomMode(e.target.value as LeafBottomMode)}
              >
                <option value="standard">Обычный профиль</option>
                <option value="freeway">FreeWay</option>
              </select>
            </label>
            <label>
              Количество сегментов
              <input
                type="number"
                value={segmentsCount}
                min={1}
                max={12}
                onChange={(e) => setSegmentsCount(Math.max(1, Number(e.target.value) || 1))}
              />
            </label>
          </div>
          <div>
            <h3>Типы сегментов (слева направо)</h3>
            <div className="segment-row">
              {normalizedSegments.map((segment, index) => (
                <label key={index}>
                  Сегмент {index + 1}
                  <select
                    value={segment}
                    onChange={(e) => setSegmentType(index, e.target.value as SegmentType)}
                  >
                    <option value="FIX">FIX (фикс)</option>
                    <option value="SL" disabled={frameSeries === '65'}>
                      SL (подвижная)
                    </option>
                  </select>
                  {segment === 'SL' ? (
                    <select
                      value={normalizedOpeningDirections[index]}
                      onChange={(e) =>
                        setOpeningDirection(index, e.target.value as OpeningDirection)
                      }
                    >
                      <option value="right">Открытие вправо →</option>
                      <option value="left">Открытие влево ←</option>
                    </select>
                  ) : null}
                </label>
              ))}
            </div>
          </div>

          <div>
            <h3>Схема рамы, Вид из помещения</h3>
            <svg viewBox="0 -24 840 384" className="scheme" role="img" aria-label="Схема рамы">
              <defs>
                <marker
                  id="arrow-end"
                  markerWidth="8"
                  markerHeight="8"
                  refX="6"
                  refY="4"
                  orient="auto"
                  markerUnits="strokeWidth"
                >
                  <path d="M0,0 L8,4 L0,8 z" className="dim-arrow" />
                </marker>
                <marker
                  id="arrow-start"
                  markerWidth="8"
                  markerHeight="8"
                  refX="2"
                  refY="4"
                  orient="auto"
                  markerUnits="strokeWidth"
                >
                  <path d="M8,0 L0,4 L8,8 z" className="dim-arrow" />
                </marker>
                <marker
                  id="opening-arrow"
                  markerWidth="10"
                  markerHeight="10"
                  refX="8"
                  refY="5"
                  orient="auto"
                  markerUnits="strokeWidth"
                >
                  <path d="M0,0 L10,5 L0,10 z" className="opening-arrow-head" />
                </marker>
              </defs>

              <line x1="20" y1="20" x2="20" y2="8" className="dim-helper" />
              <line x1="820" y1="20" x2="820" y2="8" className="dim-helper" />
              <line
                x1="20"
                y1="8"
                x2="820"
                y2="8"
                className="dim-line"
                markerStart="url(#arrow-start)"
                markerEnd="url(#arrow-end)"
              />
              <text x="420" y="6" textAnchor="middle" className="dim-text">
                {widthMm.toFixed(0)} мм
              </text>

              <line x1="20" y1="20" x2="6" y2="20" className="dim-helper" />
              <line x1="20" y1="320" x2="6" y2="320" className="dim-helper" />
              <line
                x1="6"
                y1="20"
                x2="6"
                y2="320"
                className="dim-line"
                markerStart="url(#arrow-start)"
                markerEnd="url(#arrow-end)"
              />
              <text x="3" y="170" textAnchor="middle" className="dim-text dim-text-vertical">
                {heightMm.toFixed(0)} мм
              </text>

              {normalizedSegments.map((segment, index) => {
                const x = 20 + (800 / segmentsCount) * index
                const w = 800 / segmentsCount
                return (
                  <g key={`panel-${index}`}>
                    <rect
                      x={x}
                      y={20}
                      width={w}
                      height={300}
                      className={segment === 'SL' ? 'panel-sl' : 'panel-fix'}
                    />
                    <text x={x + w / 2} y={36} textAnchor="middle" className="segment-index-text">
                      Сегмент {index + 1}
                    </text>
                    <text x={x + w / 2} y={175} textAnchor="middle" className="segment-text">
                      {segment}
                    </text>
                    <line x1={x} y1={320} x2={x} y2={338} className="dim-helper" />
                    <line x1={x + w} y1={320} x2={x + w} y2={338} className="dim-helper" />
                    <line
                      x1={x}
                      y1={338}
                      x2={x + w}
                      y2={338}
                      className="dim-line"
                      markerStart="url(#arrow-start)"
                      markerEnd="url(#arrow-end)"
                    />
                    <text x={x + w / 2} y={353} textAnchor="middle" className="dim-text">
                      {panelWidth.toFixed(0)} мм
                    </text>
                    {segment === 'SL' ? (
                      <>
                        <text x={x + w / 2} y={195} textAnchor="middle" className="leaf-size-text">
                          {panelWidth.toFixed(0)} x {segmentHeightMm.toFixed(0)} мм
                        </text>
                        <line
                          x1={
                            normalizedOpeningDirections[index] === 'left'
                              ? x + w * 0.78
                              : x + w * 0.22
                          }
                          y1={212}
                          x2={
                            normalizedOpeningDirections[index] === 'left'
                              ? x + w * 0.22
                              : x + w * 0.78
                          }
                          y2={212}
                          className="opening-arrow-line"
                          markerEnd="url(#opening-arrow)"
                        />
                      </>
                    ) : (
                      <text x={x + w / 2} y={195} textAnchor="middle" className="leaf-size-text">
                        {panelWidth.toFixed(0)} x {segmentHeightMm.toFixed(0)} мм
                      </text>
                    )}
                    {index > 0 ? (
                      <line
                        x1={x}
                        y1={20}
                        x2={x}
                        y2={320}
                        className={
                          normalizedSegments[index - 1] === 'FIX' && segment === 'FIX'
                            ? 'joint-fix-fix'
                            : 'joint-fix-sl'
                        }
                      />
                    ) : null}
                  </g>
                )
              })}
              <rect
                x="20"
                y="20"
                width="800"
                height="300"
                className="frame"
                style={{ fill: 'none' }}
              />
            </svg>
          </div>

          <div className="calc">
            <h3>Расчет</h3>
            <p>
              Система рамы: {frameSystem.title} ({frameSystem.description})
            </p>
            <ul>
              <li>Периметр рамы: {calculations.perimeter.toFixed(0)} мм</li>
              <li>Стыки FIX/FIX: {calculations.fixFixJoints}</li>
              <li>Стыки FIX/SL: {calculations.fixSlJoints}</li>
              <li>Количество SL створок: {calculations.slidingPanels}</li>
              <li>Общий периметр створок SL: {calculations.slidingPerimeter.toFixed(0)} мм</li>
              <li>
                Высота сегмента: {segmentHeightMm.toFixed(0)} мм (высота рамы {heightMm.toFixed(0)} - вычет{' '}
                {segmentHeightOffsetMm.toFixed(0)})
              </li>
              <li>Коэффициент отхода: {wastePercent}%</li>
            </ul>
          </div>

          <div className="spec">
            <h3>Состав рамы</h3>
            {readyAssemblyCompositionRows.length > 0 ? (
              <>
                <table>
                  <thead>
                    <tr>
                      <th>Зона</th>
                      <th>Готовый профиль</th>
                      <th>Кол-во, шт</th>
                      <th>Вычет, мм</th>
                      <th>Длина 1 шт, мм</th>
                      <th>Суммарная длина, мм</th>
                    </tr>
                  </thead>
                  <tbody>
                    {readyAssemblyCompositionRows.map((item) => (
                      <tr key={item.key}>
                        <td>{item.zone}</td>
                        <td>{item.article ? `${item.article} | ` : ''}{item.name}</td>
                        <td>{item.quantity}</td>
                        <td>{item.deductionLabel}</td>
                        <td>{item.lengthPerUnitMm === null ? 'см. спецификацию' : item.lengthPerUnitMm.toFixed(0)}</td>
                        <td>{item.totalLengthMm === null ? 'см. спецификацию' : item.totalLengthMm.toFixed(0)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <p className="hint">
                  В спецификации ниже показаны детали, из которых собраны эти готовые профили.
                </p>
              </>
            ) : frameComposition.length === 0 ? (
              <p className="hint">Шаблон состава пока не задан для этой конфигурации.</p>
            ) : (
              <table>
                <thead>
                  <tr>
                    <th>Артикул</th>
                    <th>Профиль</th>
                    <th>Элемент</th>
                    <th>Кол-во контуров</th>
                    <th>Длина одного контура, мм</th>
                    <th>Суммарная длина, мм</th>
                  </tr>
                </thead>
                <tbody>
                  {frameComposition.map((item) => (
                    <tr key={item.part}>
                      <td>{item.article}</td>
                      <td>{item.profileName}</td>
                      <td>{item.part}</td>
                      <td>{item.count}</td>
                      <td>{item.lengthPerPartMm.toFixed(0)}</td>
                      <td>{item.totalLengthMm.toFixed(0)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>

          <div className="spec">
            <h3>Спецификация профилей</h3>
            <div className="actions">
              <button type="button" className="danger-button" onClick={resetProjectToTemplate}>
                Сбросить к шаблону
              </button>
              <button type="button" className="add-button" onClick={saveProject}>
                Сохранить проект JSON
              </button>
              <label className="upload-button">
                Загрузить проект JSON
                <input type="file" accept=".json,application/json" onChange={loadProject} />
              </label>
              <button type="button" className="add-button" onClick={exportToExcel}>
                Экспорт в Excel
              </button>
              <button type="button" className="add-button" onClick={exportToPdf}>
                Экспорт в PDF
              </button>
            </div>
            {(() => {
              const slBottomPerLeafMm = Number(panelWidth.toFixed(0))
              return (
                <p className="hint">
                  Проверка для SL оклейка A: {calculations.slidingPanels} × {slBottomPerLeafMm} ={' '}
                  {Math.round(calculations.slidingPanels * panelWidth)} мм
                </p>
              )
            })()}
            <table>
              <thead>
                <tr>
                  <th>Артикул</th>
                  <th>Профиль</th>
                  <th>Назначение</th>
                  <th>Формула</th>
                  <th>Кол-во в сборке, шт</th>
                  <th>Суммарная длина, мм</th>
                  <th>С учетом отхода, мм</th>
                  <th>Длина хлыста, мм</th>
                  <th>Хлысты, шт</th>
                  <th>Сечение</th>
                </tr>
              </thead>
              <tbody>
                {spec.map((item) => (
                  <tr key={item.key}>
                    <td>{item.article}</td>
                    <td>{item.profile}</td>
                    <td>{item.zone}</td>
                    <td>
                      {item.isHardware
                        ? `Фурнитура, итого шт: ${item.totalLengthMm.toFixed(0)}`
                        : item.zone.startsWith('SL оклейка A (низ)')
                          ? `${calculations.slidingPanels} × ${item.partQuantity} × (${Number(panelWidth.toFixed(0))} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
                          : item.zone.startsWith('SL оклейка B/D (вертикали)')
                            ? `${calculations.slidingPanels} × 2 × ${item.partQuantity} × (${Number(segmentHeightMm.toFixed(0))} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
                            : item.zone.startsWith('SL оклейка C (верх)')
                              ? `${calculations.slidingPanels} × ${item.partQuantity} × (${Number(panelWidth.toFixed(0))} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
                            : item.zone.startsWith('Стыки FIX/FIX')
                              ? `${calculations.fixFixJoints} × ${item.partQuantity} × (${Number(segmentHeightMm.toFixed(0))} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
                              : item.zone.startsWith('Стыки FIX/SL')
                                ? `${calculations.fixSlJoints} × ${item.partQuantity} × (${Number(segmentHeightMm.toFixed(0))} - ${item.deductionMm} + ${item.additionMm}) = ${item.totalLengthMm.toFixed(0)}`
                                : `1 × ${item.totalLengthMm.toFixed(0)} = ${item.totalLengthMm.toFixed(0)}`}
                    </td>
                    <td>{item.partQuantity}</td>
                    <td>{item.isHardware ? `${item.totalLengthMm.toFixed(0)} шт` : item.totalLengthMm.toFixed(0)}</td>
                    <td>
                      {item.isHardware ? `${item.withWasteLengthMm.toFixed(0)} шт` : item.withWasteLengthMm.toFixed(0)}
                    </td>
                    <td>{item.isHardware ? '—' : item.stockLengthMm}</td>
                    <td>{item.isHardware ? `${item.barsCount} шт` : item.barsCount}</td>
                    <td>
                      {item.sectionImageDataUrl ? (
                        <button
                          type="button"
                          className="add-button"
                          onClick={() => setSectionPreviewUrl(item.sectionImageDataUrl ?? '')}
                        >
                          Открыть
                        </button>
                      ) : (
                        <span className="hint">Нет</span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="totals">
              <strong>Итого (профили):</strong> {totalSpec.totalLengthMm.toFixed(0)} мм, с отходом{' '}
              {totalSpec.withWasteLengthMm.toFixed(0)} мм, хлыстов {totalSpec.totalBars} шт.
              {totalSpec.hardwarePieces > 0 ? (
                <>
                  {' '}
                  <strong>Фурнитура:</strong> {totalSpec.hardwarePieces.toFixed(0)} шт, с отходом{' '}
                  {totalSpec.hardwareWithWaste.toFixed(0)} шт к заказу ({totalSpec.hardwareBars} ед.).
                </>
              ) : null}
            </div>
            <p className="hint">Проект сохраняется автоматически в браузере.</p>
          </div>
        </section>
      ) : activeTab === 'profiles' ? (
        <section className="content">
          <h3>База профилей</h3>
          <div className="grid">
            <label>
              Артикул
              <input
                type="text"
                value={newProfileArticle}
                onChange={(e) => setNewProfileArticle(e.target.value)}
              />
            </label>
            <label>
              Название
              <input
                type="text"
                value={newProfileName}
                onChange={(e) => setNewProfileName(e.target.value)}
              />
            </label>
            <label>
              Стандартная длина, мм
              <input
                type="number"
                value={newProfileLength}
                min={100}
                onChange={(e) => setNewProfileLength(Number(e.target.value) || 0)}
              />
            </label>
            <label>
              Сечение профиля (png/jpg)
              <input type="file" accept="image/*" onChange={handleNewProfileSectionFile} />
            </label>
          </div>
          {newProfileSectionImageDataUrl ? (
            <div className="section-thumb-row">
              <img src={newProfileSectionImageDataUrl} alt="Сечение нового профиля" className="section-thumb" />
            </div>
          ) : null}
          <button type="button" className="add-button" onClick={addProfile}>
            Добавить профиль
          </button>
          <div className="grid">
            <label>
              Поиск по артикулу или названию
              <input
                type="text"
                value={profileSearch}
                onChange={(e) => setProfileSearch(e.target.value)}
                placeholder="Например: FO-65 или термомост"
              />
            </label>
          </div>

          <table>
            <thead>
              <tr>
                <th className="col-article">Артикул</th>
                <th className="col-name">Название</th>
                <th className="col-length">Длина, мм</th>
                <th className="col-section">Сечение</th>
                <th className="col-actions">Действия</th>
              </tr>
            </thead>
            <tbody>
              {filteredProfiles.map((profile) => (
                <tr key={profile.id}>
                  <td className="col-article">
                    <input
                      type="text"
                      value={profile.article ?? ''}
                      onChange={(e) => updateProfile(profile.id, { article: e.target.value })}
                    />
                  </td>
                  <td className="col-name">
                    <input
                      type="text"
                      value={profile.name}
                      onChange={(e) => updateProfile(profile.id, { name: e.target.value })}
                    />
                  </td>
                  <td className="col-length">
                    <input
                      type="number"
                      min={100}
                      value={profile.defaultLength}
                      onChange={(e) =>
                        updateProfile(profile.id, { defaultLength: Number(e.target.value) || 0 })
                      }
                    />
                  </td>
                  <td className="col-section">
                    <div className="row-actions">
                      <label className="add-button">
                        Загрузить
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) => handleProfileSectionFile(profile.id, e)}
                          className="hidden-file-input"
                        />
                      </label>
                      {profile.sectionImageDataUrl ? (
                        <button
                          type="button"
                          className="add-button"
                          onClick={() => setSectionPreviewUrl(profile.sectionImageDataUrl ?? '')}
                        >
                          Просмотр
                        </button>
                      ) : (
                        <span className="hint">Нет</span>
                      )}
                    </div>
                  </td>
                  <td className="col-actions">
                    <div className="row-actions">
                      <button
                        type="button"
                        className="add-button"
                        onClick={() => duplicateProfile(profile.id)}
                      >
                        Дублировать
                      </button>
                      <button
                        type="button"
                        className="danger-button icon-button"
                        title="Удалить профиль"
                        aria-label="Удалить профиль"
                        onClick={() => deleteProfile(profile.id)}
                      >
                        🗑
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </section>
      ) : activeTab === 'hardware' ? (
        <section className="content">
          <h3>База фурнитуры</h3>
          <p className="hint">
            Ручки, крепёж, каретки и другие позиции без длины хлыста. Их можно добавлять в состав готовых профилей,
            шаблонов узлов и пользовательских узлов наряду с профилями (в одной строке выбирается либо профиль,
            либо фурнитура).
          </p>
          <div className="grid">
            <label>
              Артикул
              <input
                type="text"
                value={newHardwareArticle}
                onChange={(e) => setNewHardwareArticle(e.target.value)}
                placeholder="Например: HND-01"
              />
            </label>
            <label>
              Название
              <input
                type="text"
                value={newHardwareName}
                onChange={(e) => setNewHardwareName(e.target.value)}
                placeholder="Например: Ручка SL"
              />
            </label>
          </div>
          <button type="button" className="add-button" onClick={addHardware}>
            Добавить позицию
          </button>
          <div className="grid">
            <label>
              Поиск по артикулу или названию
              <input
                type="text"
                value={hardwareSearch}
                onChange={(e) => setHardwareSearch(e.target.value)}
                placeholder="Например: ручка или HND"
              />
            </label>
          </div>
          <table>
            <thead>
              <tr>
                <th className="col-article">Артикул</th>
                <th className="col-name">Название</th>
                <th className="col-actions">Действия</th>
              </tr>
            </thead>
            <tbody>
              {filteredHardware.map((item) => (
                <tr key={item.id}>
                  <td className="col-article">
                    <input
                      type="text"
                      value={item.article}
                      onChange={(e) => updateHardware(item.id, { article: e.target.value })}
                    />
                  </td>
                  <td className="col-name">
                    <input
                      type="text"
                      value={item.name}
                      onChange={(e) => updateHardware(item.id, { name: e.target.value })}
                    />
                  </td>
                  <td className="col-actions">
                    <div className="row-actions">
                      <button type="button" className="add-button" onClick={() => duplicateHardware(item.id)}>
                        Дублировать
                      </button>
                      <button
                        type="button"
                        className="danger-button icon-button"
                        title="Удалить"
                        aria-label="Удалить"
                        onClick={() => deleteHardware(item.id)}
                      >
                        🗑
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </section>
      ) : activeTab === 'ready-hardware' ? (
        <section className="content">
          <h3>Готовая фурнитура</h3>
          <p className="hint">
            Наборы фурнитуры (например: замок, каретки и т.д.) для быстрого добавления в готовые профили.
          </p>
          <div className="grid">
            <label>
              Артикул набора
              <input
                type="text"
                value={newReadyHardwareArticle}
                onChange={(e) => setNewReadyHardwareArticle(e.target.value)}
                placeholder="Например: KIT-SL-01"
              />
            </label>
            <label>
              Название набора
              <input
                type="text"
                value={newReadyHardwareName}
                onChange={(e) => setNewReadyHardwareName(e.target.value)}
                placeholder="Например: Комплект фурнитуры SL"
              />
            </label>
          </div>
          <button type="button" className="add-button" onClick={addReadyHardware}>
            Добавить набор
          </button>
          <div className="grid">
            <label>
              Поиск по артикулу или названию
              <input
                type="text"
                value={readyHardwareSearch}
                onChange={(e) => setReadyHardwareSearch(e.target.value)}
                placeholder="Например: KIT или каретка"
              />
            </label>
          </div>
          <table>
            <thead>
              <tr>
                <th className="col-article">Артикул</th>
                <th className="col-name">Название</th>
                <th className="col-actions">Действия</th>
              </tr>
            </thead>
            <tbody>
              {filteredReadyHardware.map((item) => (
                <tr key={item.id}>
                  <td className="col-article">
                    <input
                      type="text"
                      value={item.article}
                      onChange={(e) => updateReadyHardware(item.id, { article: e.target.value })}
                    />
                  </td>
                  <td className="col-name">
                    <input
                      type="text"
                      value={item.name}
                      onChange={(e) => updateReadyHardware(item.id, { name: e.target.value })}
                    />
                  </td>
                  <td className="col-actions">
                    <div className="row-actions">
                      <button
                        type="button"
                        className="add-button"
                        onClick={() => duplicateReadyHardware(item.id)}
                      >
                        Дублировать
                      </button>
                      <button
                        type="button"
                        className="danger-button icon-button"
                        title="Удалить"
                        aria-label="Удалить"
                        onClick={() => deleteReadyHardware(item.id)}
                      >
                        🗑
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </section>
      ) : activeTab === 'assemblies' ? (
        <section className="content">
          <div className="spec">
            <h3>База готовых профилей (сборок)</h3>
            <div className="grid">
              <label>
                Поиск по артикулу или названию
                <input
                  type="text"
                  value={assemblySearch}
                  onChange={(e) => setAssemblySearch(e.target.value)}
                  placeholder="Например: AS-SL-A-01 или профиль A"
                />
              </label>
            </div>
            <div className="grid">
              <label>
                Артикул сборки
                <input
                  type="text"
                  value={newAssemblyArticle}
                  onChange={(e) => setNewAssemblyArticle(e.target.value)}
                  placeholder="Например: AS-SL-A-01"
                />
              </label>
              <label>
                Название сборки
                <input
                  type="text"
                  value={newAssemblyName}
                  onChange={(e) => setNewAssemblyName(e.target.value)}
                  placeholder="Например: SL профиль A"
                />
              </label>
            </div>
            <div className="section-actions-column">
              <div className="section-actions-inline">
                <label className="add-button" title="Загрузить сечение" aria-label="Загрузить сечение">
                  Сечение сборки: загрузить
                  <input
                    type="file"
                    accept="image/*"
                    onChange={handleNewAssemblySectionFile}
                    className="hidden-file-input"
                  />
                </label>
                {newAssemblySectionImageDataUrl ? (
                  <button
                    type="button"
                    className="danger-button icon-button"
                    title="Удалить сечение"
                    aria-label="Удалить сечение"
                    onClick={() => setNewAssemblySectionImageDataUrl('')}
                  >
                    🗑
                  </button>
                ) : null}
              </div>
              {newAssemblySectionImageDataUrl ? (
                <>
                  <button
                    type="button"
                    className="add-button icon-button"
                    title="Просмотр"
                    aria-label="Просмотр"
                    onClick={() => setSectionPreviewUrl(newAssemblySectionImageDataUrl)}
                  >
                    👁
                  </button>
                </>
              ) : (
                <span className="hint">Сечение не загружено</span>
              )}
            </div>
            <p className="hint">
              В строке выберите профиль или фурнитуру (не оба). Далее: кол-во, шт, вычет и добавка в мм (для
              фурнитуры мм в спецификации не используются — считается число штук по контуру).
            </p>
            <div className="custom-parts-list">
              <div className="custom-part-row-assembly part-row-header">
                <span>№</span>
                <span>Профиль / фурнитура</span>
                <span>шт.</span>
                <span>−</span>
                <span>+</span>
                <span>Дейст.</span>
              </div>
              {newAssemblyParts.map((part, index) => (
                <div className="custom-part-row-assembly" key={part.id}>
                  <span>{index + 1}</span>
                  <div className="assembly-item-combobox">
                    {(() => {
                      const searchKey = part.id
                      const currentInput =
                        newAssemblyPartSearch[searchKey] ?? getAssemblyPartSelectedLabel(part)
                      const choices = getAssemblyPartChoices(currentInput)
                      return (
                        <>
                          <input
                            type="text"
                            list={`assembly-part-choices-${part.id}`}
                            value={currentInput}
                            placeholder="Профиль или фурнитура: начните вводить артикул/название"
                            onChange={(e) => {
                              const nextValue = e.target.value
                              setNewAssemblyPartSearch((current) => ({
                                ...current,
                                [searchKey]: nextValue,
                              }))
                              const matched = getAssemblyPartChoices(nextValue).find(
                                (choice) => choice.label === nextValue,
                              )
                              if (!matched) return
                              updateNewAssemblyPart(part.id, {
                                profileId: matched.kind === 'profile' ? matched.id : '',
                                hardwareId: matched.kind === 'hardware' ? matched.id : '',
                                readyHardwareId: matched.kind === 'readyHardware' ? matched.id : '',
                              })
                            }}
                          />
                          <datalist id={`assembly-part-choices-${part.id}`}>
                            {choices.map((choice) => (
                              <option key={`${choice.kind}-${choice.id}`} value={choice.label} />
                            ))}
                          </datalist>
                        </>
                      )
                    })()}
                  </div>
                  <input
                    type="number"
                    min={1}
                    step={1}
                    value={part.quantity}
                    className="compact-number-input"
                    placeholder="шт."
                    title="шт."
                    onChange={(e) =>
                      updateNewAssemblyPart(part.id, {
                        quantity: Math.max(1, Number(e.target.value) || 1),
                      })
                    }
                  />
                  <input
                    type="number"
                    min={0}
                    step={1}
                    value={editableZeroNumber(part.deductionMm)}
                    className="compact-number-input"
                    placeholder="−"
                    title="−"
                    onChange={(e) =>
                      updateNewAssemblyPart(part.id, {
                        deductionMm: Math.max(0, Number(e.target.value) || 0),
                      })
                    }
                  />
                  <input
                    type="number"
                    min={0}
                    step={1}
                    value={editableZeroNumber(part.additionMm)}
                    className="compact-number-input"
                    placeholder="+"
                    title="+"
                    onChange={(e) =>
                      updateNewAssemblyPart(part.id, {
                        additionMm: Math.max(0, Number(e.target.value) || 0),
                      })
                    }
                  />
                  <button
                    type="button"
                    className="danger-button icon-button"
                    title="Удалить"
                    aria-label="Удалить"
                    onClick={() => removeNewAssemblyPart(part.id)}
                    disabled={newAssemblyParts.length === 1}
                  >
                    🗑
                  </button>
                </div>
              ))}
            </div>
            <div className="actions">
              <button type="button" className="add-button" onClick={addNewAssemblyPart}>
                Добавить деталь
              </button>
              <button type="button" className="add-button" onClick={addReadyAssembly}>
                Добавить сборку
              </button>
            </div>
            <table>
              <thead>
                <tr>
                  <th className="assembly-col-article">Артикул</th>
                  <th className="assembly-col-name">Название</th>
                  <th className="assembly-col-parts">Детали</th>
                  <th className="assembly-col-actions">Сечение</th>
                </tr>
              </thead>
              <tbody>
                {filteredReadyAssemblies.map((assembly) => (
                  <tr key={assembly.id}>
                    <td className="assembly-col-article">
                      <input
                        type="text"
                        value={assembly.article}
                        onChange={(e) => updateReadyAssembly(assembly.id, { article: e.target.value })}
                      />
                    </td>
                    <td className="assembly-col-name assembly-name-cell">
                      <input
                        type="text"
                        value={assembly.name}
                        onChange={(e) => updateReadyAssembly(assembly.id, { name: e.target.value })}
                      />
                      <button
                        type="button"
                        className="danger-button icon-button"
                        title="Удалить сборку"
                        aria-label="Удалить сборку"
                        onClick={() => removeReadyAssembly(assembly.id)}
                      >
                        🗑
                      </button>
                    </td>
                    <td className="assembly-col-parts">
                      <div className="node-parts-editor">
                        <div className="assembly-part-row part-row-header">
                          <span>№</span>
                          <span>Профиль / фурнитура</span>
                          <span>шт.</span>
                          <span>−</span>
                          <span>+</span>
                          <span>Дейст.</span>
                        </div>
                        {assembly.parts.map((part, index) => (
                          <div className="assembly-part-row" key={part.id}>
                            <span>{index + 1}</span>
                            <div className="assembly-item-combobox">
                              {(() => {
                                const searchKey = `${assembly.id}::${part.id}`
                                const currentInput =
                                  readyAssemblyPartSearch[searchKey] ??
                                  getAssemblyPartSelectedLabel(part)
                                const choices = getAssemblyPartChoices(currentInput)
                                return (
                                  <>
                                    <input
                                      type="text"
                                      list={`ready-assembly-part-choices-${assembly.id}-${part.id}`}
                                      value={currentInput}
                                      placeholder="Профиль или фурнитура: начните вводить артикул/название"
                                      onChange={(e) => {
                                        const nextValue = e.target.value
                                        setReadyAssemblyPartSearch((current) => ({
                                          ...current,
                                          [searchKey]: nextValue,
                                        }))
                                        const matched = getAssemblyPartChoices(nextValue).find(
                                          (choice) => choice.label === nextValue,
                                        )
                                        if (!matched) return
                                        updatePartInReadyAssembly(assembly.id, part.id, {
                                          profileId: matched.kind === 'profile' ? matched.id : '',
                                          hardwareId: matched.kind === 'hardware' ? matched.id : '',
                                          readyHardwareId:
                                            matched.kind === 'readyHardware' ? matched.id : '',
                                        })
                                      }}
                                    />
                                    <datalist id={`ready-assembly-part-choices-${assembly.id}-${part.id}`}>
                                      {choices.map((choice) => (
                                        <option key={`${choice.kind}-${choice.id}`} value={choice.label} />
                                      ))}
                                    </datalist>
                                  </>
                                )
                              })()}
                            </div>
                            <input
                              type="number"
                              min={1}
                              step={1}
                              value={part.quantity}
                              className="compact-number-input"
                              placeholder="шт."
                              title="шт."
                              onChange={(e) =>
                                updatePartInReadyAssembly(assembly.id, part.id, {
                                  quantity: Math.max(1, Number(e.target.value) || 1),
                                })
                              }
                            />
                            <input
                              type="number"
                              min={0}
                              step={1}
                              value={editableZeroNumber(part.deductionMm)}
                              className="compact-number-input"
                              placeholder="−"
                              title="−"
                              onChange={(e) =>
                                updatePartInReadyAssembly(assembly.id, part.id, {
                                  deductionMm: Math.max(0, Number(e.target.value) || 0),
                                })
                              }
                            />
                            <input
                              type="number"
                              min={0}
                              step={1}
                              value={editableZeroNumber(part.additionMm)}
                              className="compact-number-input"
                              placeholder="+"
                              title="+"
                              onChange={(e) =>
                                updatePartInReadyAssembly(assembly.id, part.id, {
                                  additionMm: Math.max(0, Number(e.target.value) || 0),
                                })
                              }
                            />
                            <button
                              type="button"
                              className="danger-button icon-button"
                              title="Удалить"
                              aria-label="Удалить"
                              onClick={() => removePartFromReadyAssembly(assembly.id, part.id)}
                              disabled={assembly.parts.length === 1}
                            >
                              🗑
                            </button>
                          </div>
                        ))}
                        <button
                          type="button"
                          className="add-button"
                          onClick={() => addPartToReadyAssembly(assembly.id)}
                        >
                          Добавить деталь
                        </button>
                      </div>
                    </td>
                    <td className="assembly-col-actions">
                      <div className="section-actions-column">
                        <div className="section-actions-inline">
                          <label className="add-button icon-button" title="Загрузить сечение" aria-label="Загрузить сечение">
                            🖼
                            <input
                              type="file"
                              accept="image/*"
                              onChange={(e) => handleReadyAssemblySectionFile(assembly.id, e)}
                              className="hidden-file-input"
                            />
                          </label>
                          {assembly.sectionImageDataUrl ? (
                            <button
                              type="button"
                              className="danger-button icon-button"
                              title="Удалить сечение"
                              aria-label="Удалить сечение"
                              onClick={() => updateReadyAssembly(assembly.id, { sectionImageDataUrl: undefined })}
                            >
                              🗑
                            </button>
                          ) : null}
                        </div>
                        {assembly.sectionImageDataUrl ? (
                          <>
                            <button
                              type="button"
                              className="add-button icon-button"
                              title="Просмотр"
                              aria-label="Просмотр"
                              onClick={() => setSectionPreviewUrl(assembly.sectionImageDataUrl ?? '')}
                            >
                              👁
                            </button>
                          </>
                        ) : null}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      ) : activeTab === 'nodes' ? (
        <section className="content">
          <div className="spec">
            <h3>База узлов</h3>
            <div className="actions">
              <button type="button" className="add-button" onClick={addDefaultSlNodeTemplates}>
                Добавить стандартные узлы: SL / SL FW / SL Motor / SL FW Motor
              </button>
            </div>
            <div className="grid">
              <label>
                Поиск по артикулу или названию
                <input
                  type="text"
                  value={nodeTemplateSearch}
                  onChange={(e) => setNodeTemplateSearch(e.target.value)}
                  placeholder="Например: ND-SL-FIX-01 или Узел SL-FIX"
                />
              </label>
            </div>
            <div className="grid">
              <label>
                Артикул узла
                <input
                  type="text"
                  value={newNodeTemplateArticle}
                  onChange={(e) => setNewNodeTemplateArticle(e.target.value)}
                  placeholder="Например: ND-SL-FIX-01"
                />
              </label>
              <label>
                Название узла
                <input
                  type="text"
                  value={newNodeTemplateName}
                  onChange={(e) => setNewNodeTemplateName(e.target.value)}
                  placeholder="Например: Узел SL-FIX"
                />
              </label>
            </div>
            <p className="hint">
              В строке выберите профиль или фурнитуру. Далее кол-во и вычет, мм (для фурнитуры в спецификации
              считаются штуки по контуру).
            </p>
            <div className="custom-parts-list">
              <div className="custom-part-row part-row-header">
                <span>№</span>
                <span>Профиль / фурнитура</span>
                <span>шт.</span>
                <span>−</span>
                <span>Дейст.</span>
              </div>
              {newNodeTemplateParts.map((part, index) => (
                <div className="custom-part-row" key={part.id}>
                  <span>{index + 1}</span>
                  <div className="profile-select-with-search">
                    <select
                      value={part.profileId}
                      onChange={(e) =>
                        updateNewNodeTemplatePart(part.id, {
                          profileId: e.target.value,
                          hardwareId: e.target.value ? '' : part.hardwareId,
                        })
                      }
                    >
                      <option value="">Профиль не выбран</option>
                      {profiles.map((profile) => (
                        <option key={profile.id} value={profile.id}>
                          {profile.article} | {profile.name}
                        </option>
                      ))}
                    </select>
                    <select
                      value={part.hardwareId}
                      onChange={(e) =>
                        updateNewNodeTemplatePart(part.id, {
                          hardwareId: e.target.value,
                          profileId: e.target.value ? '' : part.profileId,
                        })
                      }
                    >
                      <option value="">Фурнитура не выбрана</option>
                      {hardware.map((hw) => (
                        <option key={hw.id} value={hw.id}>
                          {hw.article} | {hw.name}
                        </option>
                      ))}
                    </select>
                  </div>
                  <input
                    type="number"
                    min={1}
                    step={1}
                    value={part.quantity}
                    className="compact-number-input"
                    placeholder="шт."
                    title="шт."
                    onChange={(e) =>
                      updateNewNodeTemplatePart(part.id, {
                        quantity: Math.max(1, Number(e.target.value) || 1),
                      })
                    }
                  />
                  <input
                    type="number"
                    min={0}
                    step={1}
                    value={editableZeroNumber(part.deductionMm)}
                    className="compact-number-input"
                    placeholder="−"
                    title="−"
                    onChange={(e) =>
                      updateNewNodeTemplatePart(part.id, {
                        deductionMm: Math.max(0, Number(e.target.value) || 0),
                      })
                    }
                  />
                  <button
                    type="button"
                    className="danger-button icon-button"
                    title="Удалить"
                    aria-label="Удалить"
                    onClick={() => removeNewNodeTemplatePart(part.id)}
                    disabled={newNodeTemplateParts.length === 1}
                  >
                    🗑
                  </button>
                </div>
              ))}
            </div>
            <div className="actions">
              <button type="button" className="add-button" onClick={addNewNodeTemplatePart}>
                Добавить деталь
              </button>
              <button type="button" className="add-button" onClick={addNodeTemplate}>
                Добавить узел
              </button>
            </div>
            <table>
              <thead>
                <tr>
                  <th>Артикул</th>
                  <th>Название</th>
                  <th>Детали</th>
                  <th>Кол-во деталей</th>
                  <th>Действия</th>
                </tr>
              </thead>
              <tbody>
                {filteredNodeTemplates.map((template) => (
                  <tr key={template.id}>
                    <td>
                      <input
                        type="text"
                        value={template.article}
                        onChange={(e) => updateNodeTemplate(template.id, { article: e.target.value })}
                      />
                    </td>
                    <td>
                      <input
                        type="text"
                        value={template.name}
                        onChange={(e) => updateNodeTemplate(template.id, { name: e.target.value })}
                      />
                    </td>
                    <td>
                      <div className="node-parts-editor">
                        <div className="node-part-row part-row-header">
                          <span>№</span>
                          <span>Профиль / фурнитура</span>
                          <span>шт.</span>
                          <span>−</span>
                          <span>Дейст.</span>
                        </div>
                        {template.parts.map((part, index) => (
                          <div className="node-part-row" key={part.id}>
                            <span>{index + 1}</span>
                            <div className="profile-select-with-search">
                              <select
                                value={part.profileId}
                                onChange={(e) =>
                                  updatePartInNodeTemplate(template.id, part.id, {
                                    profileId: e.target.value,
                                    hardwareId: e.target.value ? '' : part.hardwareId,
                                  })
                                }
                              >
                                <option value="">Профиль не выбран</option>
                                {profiles.map((profile) => (
                                  <option key={profile.id} value={profile.id}>
                                    {profile.article} | {profile.name}
                                  </option>
                                ))}
                              </select>
                              <select
                                value={part.hardwareId}
                                onChange={(e) =>
                                  updatePartInNodeTemplate(template.id, part.id, {
                                    hardwareId: e.target.value,
                                    profileId: e.target.value ? '' : part.profileId,
                                  })
                                }
                              >
                                <option value="">Фурнитура не выбрана</option>
                                {hardware.map((hw) => (
                                  <option key={hw.id} value={hw.id}>
                                    {hw.article} | {hw.name}
                                  </option>
                                ))}
                              </select>
                            </div>
                            <input
                              type="number"
                              min={1}
                              step={1}
                              value={part.quantity}
                              className="compact-number-input"
                              placeholder="шт."
                              title="шт."
                              onChange={(e) =>
                                updatePartInNodeTemplate(template.id, part.id, {
                                  quantity: Math.max(1, Number(e.target.value) || 1),
                                })
                              }
                            />
                            <input
                              type="number"
                              min={0}
                              step={1}
                              value={editableZeroNumber(part.deductionMm)}
                              className="compact-number-input"
                              placeholder="−"
                              title="−"
                              onChange={(e) =>
                                updatePartInNodeTemplate(template.id, part.id, {
                                  deductionMm: Math.max(0, Number(e.target.value) || 0),
                                })
                              }
                            />
                            <button
                              type="button"
                              className="danger-button icon-button"
                              title="Удалить"
                              aria-label="Удалить"
                              onClick={() => removePartFromNodeTemplate(template.id, part.id)}
                              disabled={template.parts.length === 1}
                            >
                              🗑
                            </button>
                          </div>
                        ))}
                        <button
                          type="button"
                          className="add-button"
                          onClick={() => addPartToNodeTemplate(template.id)}
                        >
                          Добавить деталь
                        </button>
                      </div>
                    </td>
                    <td>{template.parts.length}</td>
                    <td>
                      <button
                        type="button"
                        className="danger-button icon-button"
                        title="Удалить узел"
                        aria-label="Удалить узел"
                        onClick={() => removeNodeTemplate(template.id)}
                      >
                        🗑
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      ) : (

        <section className="content">
          <h3>Настройки</h3>
          <p className="hint">
            <strong>Где лежат ваши данные.</strong> База хранится на сервере в файле{' '}
            <code>server-data/project.json</code> в папке сайта. При открытии страницы проект загружается с
            сервера, а изменения сохраняются автоматически в этот файл.
          </p>
          <div className="grid">
            <label>
              Автосохранение
              <input type="text" value="Включено (server-data/project.json)" readOnly />
            </label>
            <label>
              Отход, %
              <input
                type="number"
                value={wastePercent}
                min={0}
                max={50}
                onChange={(e) => setWastePercent(Math.max(0, Number(e.target.value) || 0))}
              />
            </label>
            <label>
              Серия 65: вычет высоты сегмента, мм
              <input
                type="number"
                value={editableZeroNumber(segmentHeightOffset65Mm)}
                min={0}
                onChange={(e) => setSegmentHeightOffset65Mm(Math.max(0, Number(e.target.value) || 0))}
              />
            </label>
            <label>
              Серия 80: вычет высоты сегмента, мм
              <input
                type="number"
                value={editableZeroNumber(segmentHeightOffset80Mm)}
                min={0}
                onChange={(e) => setSegmentHeightOffset80Mm(Math.max(0, Number(e.target.value) || 0))}
              />
            </label>
          </div>

          <div className="actions">
            <button type="button" className="add-button" onClick={resetProjectToTemplate}>
              Сбросить проект к шаблону
            </button>
          </div>
          <p className="hint">Серверная база: {serverSyncStatus}</p>

          <div className="spec">
            <h3>Назначение узлов из базы</h3>
            {!activeSlBottomBindingId ? (
              <p className="warning-text">
                Для текущего режима нижнего профиля створки не назначен узел/сборка `SL оклейка A (низ)`.
              </p>
            ) : null}
            <div className="grid">
              <label>
                Однотре́ковый профиль рамы
                <select
                  value={systemNodeTemplateBinding.frameSingleTrackAssembly}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({
                      ...s,
                      frameSingleTrackAssembly: e.target.value,
                    }))
                  }
                >
                  <option value="">Готовый профиль не выбран</option>
                  {readyAssemblies.map((assembly) => (
                    <option key={assembly.id} value={assembly.id}>
                      {assembly.article ? `${assembly.article} | ` : ''}
                      {assembly.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Двухтрековый профиль рамы
                <select
                  value={systemNodeTemplateBinding.frameDoubleTrackAssembly}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({
                      ...s,
                      frameDoubleTrackAssembly: e.target.value,
                    }))
                  }
                >
                  <option value="">Готовый профиль не выбран</option>
                  {readyAssemblies.map((assembly) => (
                    <option key={assembly.id} value={assembly.id}>
                      {assembly.article ? `${assembly.article} | ` : ''}
                      {assembly.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Трехтрековый профиль рамы
                <select
                  value={systemNodeTemplateBinding.frameTripleTrackAssembly}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({
                      ...s,
                      frameTripleTrackAssembly: e.target.value,
                    }))
                  }
                >
                  <option value="">Готовый профиль не выбран</option>
                  {readyAssemblies.map((assembly) => (
                    <option key={assembly.id} value={assembly.id}>
                      {assembly.article ? `${assembly.article} | ` : ''}
                      {assembly.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Четырехтрековый профиль рамы
                <select
                  value={systemNodeTemplateBinding.frameQuadrupleTrackAssembly}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({
                      ...s,
                      frameQuadrupleTrackAssembly: e.target.value,
                    }))
                  }
                >
                  <option value="">Готовый профиль не выбран</option>
                  {readyAssemblies.map((assembly) => (
                    <option key={assembly.id} value={assembly.id}>
                      {assembly.article ? `${assembly.article} | ` : ''}
                      {assembly.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Стык FIX/FIX
                <select
                  value={systemNodeTemplateBinding.fixFix}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, fixFix: e.target.value }))
                  }
                >
                  <option value="">Узел не выбран</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Стык FIX/SL
                <select
                  value={systemNodeTemplateBinding.fixSl}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, fixSl: e.target.value }))
                  }
                >
                  <option value="">Узел не выбран</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                SL оклейка A (низ)
                <select
                  value={systemNodeTemplateBinding.slLeafBottomA}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, slLeafBottomA: e.target.value }))
                  }
                >
                  <option value="">Узел/сборка не выбраны</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                  {readyAssemblies.map((assembly) => (
                    <option key={assembly.id} value={assembly.id}>
                      {assembly.article ? `${assembly.article} | ` : ''}
                      {assembly.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                SL оклейка A (низ, FreeWay)
                <select
                  value={systemNodeTemplateBinding.slLeafBottomAFreeway}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, slLeafBottomAFreeway: e.target.value }))
                  }
                >
                  <option value="">Узел/сборка не выбраны</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                  {readyAssemblies.map((assembly) => (
                    <option key={assembly.id} value={assembly.id}>
                      {assembly.article ? `${assembly.article} | ` : ''}
                      {assembly.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                SL оклейка B/D (вертикали)
                <select
                  value={systemNodeTemplateBinding.slLeafVerticalBd}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, slLeafVerticalBd: e.target.value }))
                  }
                >
                  <option value="">Узел/сборка не выбраны</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                  {readyAssemblies.map((assembly) => (
                    <option key={assembly.id} value={assembly.id}>
                      {assembly.article ? `${assembly.article} | ` : ''}
                      {assembly.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                SL оклейка C (верх)
                <select
                  value={systemNodeTemplateBinding.slLeafTopC}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, slLeafTopC: e.target.value }))
                  }
                >
                  <option value="">Узел/сборка не выбраны</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                  {readyAssemblies.map((assembly) => (
                    <option key={assembly.id} value={assembly.id}>
                      {assembly.article ? `${assembly.article} | ` : ''}
                      {assembly.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                SL рабочая сторона → рама
                <select
                  value={systemNodeTemplateBinding.slSideToFrame}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, slSideToFrame: e.target.value }))
                  }
                >
                  <option value="">Узел не выбран</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                SL рабочая сторона → FIX
                <select
                  value={systemNodeTemplateBinding.slSideToFix}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, slSideToFix: e.target.value }))
                  }
                >
                  <option value="">Узел не выбран</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                SL рабочая сторона → SL
                <select
                  value={systemNodeTemplateBinding.slSideToSl}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, slSideToSl: e.target.value }))
                  }
                >
                  <option value="">Узел не выбран</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                SL рабочая сторона → пенал
                <select
                  value={systemNodeTemplateBinding.slSideToPenal}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, slSideToPenal: e.target.value }))
                  }
                >
                  <option value="">Узел не выбран</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                SL рабочая сторона → угол
                <select
                  value={systemNodeTemplateBinding.slSideToCorner}
                  onChange={(e) =>
                    setSystemNodeTemplateBinding((s) => ({ ...s, slSideToCorner: e.target.value }))
                  }
                >
                  <option value="">Узел не выбран</option>
                  {nodeTemplates.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.article ? `${template.article} | ` : ''}
                      {template.name}
                    </option>
                  ))}
                </select>
              </label>
            </div>
          </div>

          <div className="spec">
            <h3>Пользовательские узлы</h3>
            <div className="grid">
              <label>
                Название узла
                <input
                  type="text"
                  value={newNodeName}
                  onChange={(e) => setNewNodeName(e.target.value)}
                />
              </label>
              <label>
                Формула длины
                <select
                  value={newNodeFormula}
                  onChange={(e) => setNewNodeFormula(e.target.value as CustomNodeFormula)}
                >
                  {Object.entries(formulaLabels).map(([key, label]) => (
                    <option key={key} value={key}>
                      {label}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Режим количества
                <select
                  value={newNodeQuantityMode}
                  onChange={(e) =>
                    setNewNodeQuantityMode(e.target.value as CustomNodeQuantityMode)
                  }
                >
                  {Object.entries(quantityModeLabels).map(([key, label]) => (
                    <option key={key} value={key}>
                      {label}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Значение (для фиксированного)
                <input
                  type="number"
                  min={1}
                  value={newNodeQuantityValue}
                  onChange={(e) => setNewNodeQuantityValue(Math.max(1, Number(e.target.value) || 1))}
                />
              </label>
            </div>
            <h4>Состав узла</h4>
            <p className="hint">
              В строке выберите профиль или фурнитуру. Для фурнитуры в спецификации считается число штук: кол-во
              в строке × множитель по формуле узла.
            </p>
            <div className="custom-parts-list">
              <div className="custom-part-row part-row-header">
                <span>№</span>
                <span>Профиль / фурнитура</span>
                <span>шт.</span>
                <span>−</span>
                <span>Дейст.</span>
              </div>
              {newNodeParts.map((part, index) => (
                <div className="custom-part-row" key={part.id}>
                  <span>{index + 1}</span>
                  <div className="profile-select-with-search">
                    <select
                      value={part.profileId}
                      onChange={(e) =>
                        updateNewNodePart(part.id, {
                          profileId: e.target.value,
                          hardwareId: e.target.value ? '' : part.hardwareId,
                        })
                      }
                    >
                      <option value="">Профиль не выбран</option>
                      {profiles.map((profile) => (
                        <option key={profile.id} value={profile.id}>
                          {profile.article} | {profile.name}
                        </option>
                      ))}
                    </select>
                    <select
                      value={part.hardwareId}
                      onChange={(e) =>
                        updateNewNodePart(part.id, {
                          hardwareId: e.target.value,
                          profileId: e.target.value ? '' : part.profileId,
                        })
                      }
                    >
                      <option value="">Фурнитура не выбрана</option>
                      {hardware.map((hw) => (
                        <option key={hw.id} value={hw.id}>
                          {hw.article} | {hw.name}
                        </option>
                      ))}
                    </select>
                  </div>
                  <input
                    type="number"
                    min={1}
                    step={1}
                    value={part.quantity}
                    className="compact-number-input"
                    placeholder="шт."
                    title="шт."
                    onChange={(e) =>
                      updateNewNodePart(part.id, {
                        quantity: Math.max(1, Number(e.target.value) || 1),
                      })
                    }
                  />
                  <input
                    type="number"
                    min={0}
                    step={1}
                    value={editableZeroNumber(part.deductionMm)}
                    className="compact-number-input"
                    placeholder="−"
                    title="−"
                    onChange={(e) =>
                      updateNewNodePart(part.id, {
                        deductionMm: Math.max(0, Number(e.target.value) || 0),
                      })
                    }
                  />
                  <button
                    type="button"
                    className="danger-button icon-button"
                    title="Удалить"
                    aria-label="Удалить"
                    onClick={() => removeNewNodePart(part.id)}
                    disabled={newNodeParts.length === 1}
                  >
                    🗑
                  </button>
                </div>
              ))}
            </div>
            <div className="actions">
              <button type="button" className="add-button" onClick={addNewNodePart}>
                Добавить часть
              </button>
              <button type="button" className="add-button" onClick={addCustomNode}>
                Добавить узел
              </button>
            </div>
            <table>
              <thead>
                <tr>
                  <th>Узел</th>
                  <th>Состав</th>
                  <th>Формула</th>
                  <th>Количество</th>
                  <th>Действия</th>
                </tr>
              </thead>
              <tbody>
                {customNodes.map((node) => {
                  return (
                    <tr key={node.id}>
                      <td>
                        <input
                          type="text"
                          value={node.name}
                          onChange={(e) => updateCustomNode(node.id, { name: e.target.value })}
                        />
                      </td>
                      <td>
                        <div className="node-parts-editor">
                        <div className="node-part-row part-row-header">
                          <span>№</span>
                          <span>Профиль / фурнитура</span>
                          <span>шт.</span>
                          <span>−</span>
                          <span>Дейст.</span>
                        </div>
                          {node.parts.map((part, partIndex) => (
                            <div className="node-part-row" key={part.id}>
                              <span>{partIndex + 1}</span>
                              <div className="profile-select-with-search">
                                <select
                                  value={part.profileId}
                                  onChange={(e) =>
                                    updatePartInCustomNode(node.id, part.id, {
                                      profileId: e.target.value,
                                      hardwareId: e.target.value ? '' : part.hardwareId,
                                    })
                                  }
                                >
                                  <option value="">Профиль не выбран</option>
                                  {profiles.map((profileOption) => (
                                    <option key={profileOption.id} value={profileOption.id}>
                                      {profileOption.article} | {profileOption.name}
                                    </option>
                                  ))}
                                </select>
                                <select
                                  value={part.hardwareId}
                                  onChange={(e) =>
                                    updatePartInCustomNode(node.id, part.id, {
                                      hardwareId: e.target.value,
                                      profileId: e.target.value ? '' : part.profileId,
                                    })
                                  }
                                >
                                  <option value="">Фурнитура не выбрана</option>
                                  {hardware.map((hw) => (
                                    <option key={hw.id} value={hw.id}>
                                      {hw.article} | {hw.name}
                                    </option>
                                  ))}
                                </select>
                              </div>
                              <input
                                type="number"
                                min={1}
                                step={1}
                                value={part.quantity}
                                className="compact-number-input"
                                placeholder="шт."
                                title="шт."
                                onChange={(e) =>
                                  updatePartInCustomNode(node.id, part.id, {
                                    quantity: Math.max(1, Number(e.target.value) || 1),
                                  })
                                }
                              />
                              <input
                                type="number"
                                min={0}
                                step={1}
                                value={editableZeroNumber(part.deductionMm)}
                                className="compact-number-input"
                                placeholder="−"
                                title="−"
                                onChange={(e) =>
                                  updatePartInCustomNode(node.id, part.id, {
                                    deductionMm: Math.max(0, Number(e.target.value) || 0),
                                  })
                                }
                              />
                              <button
                                type="button"
                                className="danger-button icon-button"
                                title="Удалить"
                                aria-label="Удалить"
                                onClick={() => removePartFromCustomNode(node.id, part.id)}
                                disabled={node.parts.length === 1}
                              >
                                🗑
                              </button>
                            </div>
                          ))}
                          <button
                            type="button"
                            className="add-button"
                            onClick={() => addPartToCustomNode(node.id)}
                          >
                            Добавить часть
                          </button>
                        </div>
                      </td>
                      <td>
                        <select
                          value={node.formula}
                          onChange={(e) =>
                            updateCustomNode(node.id, { formula: e.target.value as CustomNodeFormula })
                          }
                        >
                          {Object.entries(formulaLabels).map(([key, label]) => (
                            <option key={key} value={key}>
                              {label}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td>
                        <select
                          value={node.quantityMode}
                          onChange={(e) =>
                            updateCustomNode(node.id, {
                              quantityMode: e.target.value as CustomNodeQuantityMode,
                            })
                          }
                        >
                          {Object.entries(quantityModeLabels).map(([key, label]) => (
                            <option key={key} value={key}>
                              {label}
                            </option>
                          ))}
                        </select>
                        {node.quantityMode === 'fixed' ? (
                          <input
                            type="number"
                            min={1}
                            value={node.quantityValue}
                            onChange={(e) =>
                              updateCustomNode(node.id, {
                                quantityValue: Math.max(1, Number(e.target.value) || 1),
                              })
                            }
                          />
                        ) : (
                          <div className="hint">{quantityModeLabels[node.quantityMode]}</div>
                        )}
                      </td>
                      <td>
                        <button
                          type="button"
                          className="add-button"
                          onClick={() => duplicateCustomNode(node.id)}
                        >
                          Дублировать
                        </button>
                        <button
                          type="button"
                          className="danger-button icon-button"
                          title="Удалить узел"
                          aria-label="Удалить узел"
                          onClick={() => removeCustomNode(node.id)}
                        >
                          🗑
                        </button>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>

          <div className="spec">
            <h3>Карта узлов и профилей</h3>
            <table>
              <thead>
                <tr>
                  <th>Узел</th>
                  <th>Артикул</th>
                  <th>Профиль</th>
                  <th>Статус</th>
                </tr>
              </thead>
              <tbody>
                {nodeProfileMap.map((row) => (
                  <tr key={row.node}>
                    <td>{row.node}</td>
                    <td>{row.article}</td>
                    <td>{row.name}</td>
                    <td>{row.status}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="hint">
            Очистка локальных данных удалит автосохраненный проект из браузера на этом устройстве.
          </p>
        </section>
      )}
      {sectionPreviewUrl ? (
        <div className="section-modal-backdrop" onClick={() => setSectionPreviewUrl('')}>
          <div className="section-modal" onClick={(e) => e.stopPropagation()}>
            <img src={sectionPreviewUrl} alt="Сечение профиля" className="section-modal-image" />
            <button
              type="button"
              className="danger-button"
              onClick={() => setSectionPreviewUrl('')}
            >
              Закрыть
            </button>
          </div>
        </div>
      ) : null}
    </main>
  )
}

export default App
