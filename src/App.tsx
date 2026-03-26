import { type FormEvent, useCallback, useEffect, useMemo, useState } from 'react'
import { EmbeddedViewSdk } from 'forma-embedded-view-sdk'
import {
  Alert,
  Box,
  Button,
  CircularProgress,
  MenuItem,
  Modal,
  Select,
  Stack,
  TextField,
  Typography,
  type AlertColor,
} from '@weave-mui/material'
import './App.css'

const HOST_TIMEOUT_MS = 2000
const REQUEST_TIMEOUT_MS = 10000
const INTERACTIVE_REQUEST_TIMEOUT_MS = 300000
const HEX_COLOR_PATTERN = /^#[0-9A-Fa-f]{6}$/
const BUILT_IN_FUNCTION_IDS = new Set([
  'residential',
  'commercial',
  'unspecified',
])

type SettingsResponse = Awaited<ReturnType<EmbeddedViewSdk['settings']['get']>>
type BuildingFunction = SettingsResponse['buildingFunctions'][number]
type BuildingCreationInput = {
  name: string
  functionId: string
  width: number
  depth: number
  floorCount: number
  floorHeight: number
}
type Notice = {
  severity: AlertColor
  text: string
}

function withTimeout<T>(
  promise: Promise<T>,
  label: string,
  timeoutMs = REQUEST_TIMEOUT_MS,
): Promise<T> {
  return new Promise((resolve, reject) => {
    const timeoutId = window.setTimeout(() => {
      reject(
        new Error(
          `${label} timed out after ${timeoutMs}ms. Make sure this page is running inside Autodesk Forma.`,
        ),
      )
    }, timeoutMs)

    promise
      .then((value) => {
        window.clearTimeout(timeoutId)
        resolve(value)
      })
      .catch((error: unknown) => {
        window.clearTimeout(timeoutId)
        reject(error)
      })
  })
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message
  }

  return 'Unknown error'
}

function readSafe(getter: () => string): string | null {
  try {
    return getter()
  } catch {
    return null
  }
}

function buildBuildingFunctionPayload(name: string, color: string) {
  const trimmedName = name.trim()
  const normalizedColor = color.trim().toUpperCase()

  if (!trimmedName) {
    throw new Error('Name is required.')
  }

  if (!normalizedColor) {
    return { name: trimmedName }
  }

  if (!HEX_COLOR_PATTERN.test(normalizedColor)) {
    throw new Error('Color must be a 6-digit hex value like #FF5733.')
  }

  return {
    name: trimmedName,
    color: normalizedColor as `#${string}`,
  }
}

function getColorPreview(color: string) {
  const normalizedColor = color.trim().toUpperCase()

  if (!normalizedColor || !HEX_COLOR_PATTERN.test(normalizedColor)) {
    return null
  }

  return normalizedColor
}

function isBuiltInFunction(buildingFunction: BuildingFunction) {
  return BUILT_IN_FUNCTION_IDS.has(buildingFunction.id)
}

function parsePositiveNumber(value: string, label: string) {
  const parsedValue = Number(value)

  if (!Number.isFinite(parsedValue) || parsedValue <= 0) {
    throw new Error(`${label} must be a positive number.`)
  }

  return parsedValue
}

function parsePositiveInteger(value: string, label: string) {
  const parsedValue = Number(value)

  if (!Number.isInteger(parsedValue) || parsedValue <= 0) {
    throw new Error(`${label} must be a positive whole number.`)
  }

  return parsedValue
}

function buildBuildingCreationInput(input: {
  name: string
  functionId: string
  width: string
  depth: string
  floorCount: string
  floorHeight: string
}): BuildingCreationInput {
  const trimmedName = input.name.trim()

  if (!trimmedName) {
    throw new Error('Building name is required.')
  }

  if (!input.functionId) {
    throw new Error('Select a building function before creating a building.')
  }

  return {
    name: trimmedName,
    functionId: input.functionId,
    width: parsePositiveNumber(input.width, 'Width'),
    depth: parsePositiveNumber(input.depth, 'Depth'),
    floorCount: parsePositiveInteger(input.floorCount, 'Floor count'),
    floorHeight: parsePositiveNumber(input.floorHeight, 'Floor height'),
  }
}

function buildFloorStackRequest(input: BuildingCreationInput) {
  const halfWidth = input.width / 2
  const halfDepth = input.depth / 2
  const planId = 'grounded-rectangle-plan'

  return {
    floors: Array.from({ length: input.floorCount }, () => ({
      planId,
      height: input.floorHeight,
    })),
    plans: [
      {
        id: planId,
        vertices: [
          { id: 'v1', x: -halfWidth, y: -halfDepth },
          { id: 'v2', x: halfWidth, y: -halfDepth },
          { id: 'v3', x: halfWidth, y: halfDepth },
          { id: 'v4', x: -halfWidth, y: halfDepth },
        ],
        units: [
          {
            polygon: ['v1', 'v2', 'v3', 'v4'],
            program: 'LIVING_UNIT' as const,
            functionId: input.functionId,
            holes: [],
          },
        ],
      },
    ],
  }
}

function buildTranslationTransform(x: number, y: number, z: number) {
  return [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, x, y, z, 1]
}

function formatCoordinate(value: number) {
  return value.toFixed(2)
}

function App() {
  const [sdk, setSdk] = useState<EmbeddedViewSdk | null>(null)
  const [hostStatus, setHostStatus] = useState('Checking Autodesk Forma connection...')
  const [hostError, setHostError] = useState<string | null>(null)
  const [projectId, setProjectId] = useState<string | null>(null)
  const [canEdit, setCanEdit] = useState(false)
  const [settingsLoading, setSettingsLoading] = useState(true)
  const [settingsError, setSettingsError] = useState<string | null>(null)
  const [busyAction, setBusyAction] = useState<string | null>(null)
  const [buildingFunctions, setBuildingFunctions] = useState<BuildingFunction[]>([])
  const [newName, setNewName] = useState('Settings API Test')
  const [newColor, setNewColor] = useState('#4F46E5')
  const [editingId, setEditingId] = useState<string | null>(null)
  const [editName, setEditName] = useState('')
  const [editColor, setEditColor] = useState('')
  const [createFunctionId, setCreateFunctionId] = useState('')
  const [createBuildingName, setCreateBuildingName] = useState('2.5D Function Test Building')
  const [createWidth, setCreateWidth] = useState('24')
  const [createDepth, setCreateDepth] = useState('18')
  const [createFloorCount, setCreateFloorCount] = useState('4')
  const [createFloorHeight, setCreateFloorHeight] = useState('3.2')
  const [creationStatus, setCreationStatus] = useState<string | null>(null)
  const [pendingDeleteFunction, setPendingDeleteFunction] =
    useState<BuildingFunction | null>(null)
  const [notice, setNotice] = useState<Notice | null>(null)

  const hostOrigin = new URLSearchParams(window.location.search).get('origin')
  const isBusy = busyAction !== null
  const canMutate = sdk !== null && canEdit && !isBusy
  const canCreateBuildings = canMutate && buildingFunctions.length > 0
  const selectedBuildingFunction = useMemo(
    () =>
      buildingFunctions.find((buildingFunction) => buildingFunction.id === editingId) ??
      null,
    [buildingFunctions, editingId],
  )
  const selectedCreationFunction = useMemo(
    () =>
      buildingFunctions.find((buildingFunction) => buildingFunction.id === createFunctionId) ??
      null,
    [buildingFunctions, createFunctionId],
  )
  const newColorPreview = getColorPreview(newColor)
  const editColorPreview = getColorPreview(editColor)
  const hasCustomFunctions = buildingFunctions.some(
    (buildingFunction) => !isBuiltInFunction(buildingFunction),
  )
  const sectionSx = {
    borderColor: 'divider',
    borderTop: 1,
    paddingTop: 3,
  }

  const hostBanner = useMemo(() => {
    if (!hostOrigin) {
      return {
        severity: 'info' as const,
        title: 'Preview mode',
        text: 'Open this page inside Autodesk Forma to load and manage project-level building functions and create 2.5D buildings.',
      }
    }

    if (hostError) {
      return {
        severity: 'error' as const,
        title: hostStatus,
        text: hostError,
      }
    }

    if (!sdk) {
      return {
        severity: 'info' as const,
        title: hostStatus,
        text: 'Connecting to the Autodesk Forma embedded view host.',
      }
    }

    if (canEdit) {
      return {
        severity: 'success' as const,
        title: 'Connected with edit access',
        text: 'You can manage custom building functions and create 2.5D buildings.',
      }
    }

    return {
      severity: 'warning' as const,
      title: 'Connected in read-only mode',
      text: 'You can review building functions, but project mutations and 2.5D building creation are disabled.',
    }
  }, [canEdit, hostError, hostOrigin, hostStatus, sdk])

  const applySettingsResponse = useCallback((response: SettingsResponse) => {
    setBuildingFunctions(response.buildingFunctions)
    setSettingsError(null)
  }, [])

  const clearEditing = useCallback(() => {
    setEditingId(null)
    setEditName('')
    setEditColor('')
  }, [])

  const refreshSettings = useCallback(
    async (
      instance: EmbeddedViewSdk,
      options: {
        reason?: string
        announceSuccess?: boolean
      } = {},
    ) => {
      const { reason = 'Refresh', announceSuccess = false } = options

      setSettingsLoading(true)
      setSettingsError(null)

      try {
        const response = await withTimeout(
          instance.settings.get(),
          'Forma.settings.get()',
        )

        applySettingsResponse(response)

        if (announceSuccess) {
          setNotice({
            severity: 'success',
            text: `Loaded ${response.buildingFunctions.length} building functions.`,
          })
        }
      } catch (error) {
        const message = getErrorMessage(error)
        setSettingsError(message)

        if (reason !== 'Initial load') {
          setNotice({
            severity: 'error',
            text: `${reason} failed: ${message}`,
          })
        }
      } finally {
        setSettingsLoading(false)
      }
    },
    [applySettingsResponse],
  )

  useEffect(() => {
    let cancelled = false

    async function initialize() {
      if (!hostOrigin) {
        setHostStatus('Preview mode')
        setHostError(null)
        setSettingsLoading(false)
        return
      }

      let instance: EmbeddedViewSdk

      try {
        instance = new EmbeddedViewSdk()
      } catch (error) {
        setHostStatus('SDK bootstrap failed')
        setHostError(getErrorMessage(error))
        setSettingsLoading(false)
        return
      }

      try {
        await withTimeout(instance.ping(), 'Forma host handshake', HOST_TIMEOUT_MS)

        if (cancelled) {
          return
        }

        setSdk(instance)
        setHostStatus('Connected to Autodesk Forma')
        setHostError(null)
        setProjectId(readSafe(() => instance.getProjectId()))

        const editAccess = await withTimeout(
          instance.getCanEdit(),
          'Forma.getCanEdit()',
        )

        if (cancelled) {
          return
        }

        setCanEdit(editAccess)
        await refreshSettings(instance, {
          reason: 'Initial load',
        })
      } catch (error) {
        if (cancelled) {
          return
        }

        setHostStatus('Unable to connect to Autodesk Forma')
        setHostError(getErrorMessage(error))
        setSettingsLoading(false)
      }
    }

    void initialize()

    return () => {
      cancelled = true
    }
  }, [hostOrigin, refreshSettings])

  useEffect(() => {
    if (editingId !== null && selectedBuildingFunction === null) {
      clearEditing()
    }
  }, [clearEditing, editingId, selectedBuildingFunction])

  useEffect(() => {
    if (buildingFunctions.length === 0) {
      if (createFunctionId !== '') {
        setCreateFunctionId('')
      }
      return
    }

    const hasSelectedFunction = buildingFunctions.some(
      (buildingFunction) => buildingFunction.id === createFunctionId,
    )

    if (!hasSelectedFunction) {
      setCreateFunctionId(buildingFunctions[0].id)
    }
  }, [buildingFunctions, createFunctionId])

  async function handleRefresh() {
    if (!sdk) {
      return
    }

    setBusyAction('refresh')
    setNotice(null)
    await refreshSettings(sdk, {
      reason: 'Refresh',
      announceSuccess: true,
    })
    setBusyAction(null)
  }

  async function handleAdd(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    if (!sdk || !canEdit) {
      return
    }

    const previousIds = new Set(buildingFunctions.map((buildingFunction) => buildingFunction.id))

    try {
      const payload = buildBuildingFunctionPayload(newName, newColor)
      setBusyAction('add')
      setNotice(null)
      setSettingsError(null)

      const response = await withTimeout(
        sdk.settings.buildingFunctions.add(payload),
        'Forma.settings.buildingFunctions.add()',
      )

      applySettingsResponse(response)

      const createdBuildingFunction = response.buildingFunctions.find(
        (buildingFunction) => !previousIds.has(buildingFunction.id),
      )

      if (createdBuildingFunction && !isBuiltInFunction(createdBuildingFunction)) {
        setEditingId(createdBuildingFunction.id)
        setEditName(createdBuildingFunction.name)
        setEditColor(createdBuildingFunction.color ?? '')
        setCreateFunctionId(createdBuildingFunction.id)
      }

      setNotice({
        severity: 'success',
        text: `Created "${payload.name}".`,
      })
    } catch (error) {
      const message = getErrorMessage(error)
      setSettingsError(message)
      setNotice({
        severity: 'error',
        text: `Could not create the building function: ${message}`,
      })
    } finally {
      setBusyAction(null)
    }
  }

  function handleSelectForUpdate(buildingFunction: BuildingFunction) {
    setEditingId(buildingFunction.id)
    setEditName(buildingFunction.name)
    setEditColor(buildingFunction.color ?? '')
  }

  async function handleUpdate(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    if (!sdk || !canEdit || !editingId) {
      return
    }

    try {
      const payload = buildBuildingFunctionPayload(editName, editColor)
      setBusyAction('update')
      setNotice(null)
      setSettingsError(null)

      const response = await withTimeout(
        sdk.settings.buildingFunctions.update({
          id: editingId,
          ...payload,
        }),
        'Forma.settings.buildingFunctions.update()',
      )

      applySettingsResponse(response)

      const updatedBuildingFunction = response.buildingFunctions.find(
        (buildingFunction) => buildingFunction.id === editingId,
      )

      if (updatedBuildingFunction) {
        setEditName(updatedBuildingFunction.name)
        setEditColor(updatedBuildingFunction.color ?? '')
      }

      setNotice({
        severity: 'success',
        text: `Updated "${payload.name}".`,
      })
    } catch (error) {
      const message = getErrorMessage(error)
      setSettingsError(message)
      setNotice({
        severity: 'error',
        text: `Could not update the building function: ${message}`,
      })
    } finally {
      setBusyAction(null)
    }
  }

  async function handleCreateBuilding(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    if (!sdk || !canEdit) {
      return
    }

    setCreationStatus(null)
    setNotice(null)

    let payload: BuildingCreationInput

    try {
      payload = buildBuildingCreationInput({
        name: createBuildingName,
        functionId: createFunctionId,
        width: createWidth,
        depth: createDepth,
        floorCount: createFloorCount,
        floorHeight: createFloorHeight,
      })
    } catch (error) {
      setNotice({
        severity: 'error',
        text: getErrorMessage(error),
      })
      return
    }

    const selectedFunctionLabel = selectedCreationFunction?.name ?? payload.functionId

    try {
      setBusyAction('create-building')
      setCreationStatus('Creating the 2.5D building element...')

      const { urn } = await withTimeout(
        sdk.elements.floorStack.createFromFloors(buildFloorStackRequest(payload)),
        'Forma.elements.floorStack.createFromFloors()',
      )

      setCreationStatus('Building element created. Click a point in the scene to place it.')

      const point = await withTimeout(
        sdk.designTool.getPoint(),
        'Forma.designTool.getPoint()',
        INTERACTIVE_REQUEST_TIMEOUT_MS,
      )

      if (!point) {
        setCreationStatus(null)
        setNotice({
          severity: 'info',
          text: `Placement cancelled for "${payload.name}".`,
        })
        return
      }

      const coordinatesLabel = `(${formatCoordinate(point.x)}, ${formatCoordinate(point.y)}, ${formatCoordinate(point.z)})`

      await withTimeout(
        sdk.proposal.addElement({
          urn,
          name: payload.name,
          transform: buildTranslationTransform(point.x, point.y, point.z),
        }),
        'Forma.proposal.addElement()',
      )

      setCreationStatus(null)
      setNotice({
        severity: 'success',
        text: `Created "${payload.name}" at ${coordinatesLabel} with function "${selectedFunctionLabel}".`,
      })
    } catch (error) {
      setCreationStatus(null)
      setNotice({
        severity: 'error',
        text: `Could not create the 2.5D building: ${getErrorMessage(error)}`,
      })
    } finally {
      setBusyAction(null)
    }
  }

  function handleRequestDelete(buildingFunction: BuildingFunction) {
    if (!canMutate || isBuiltInFunction(buildingFunction)) {
      return
    }

    setPendingDeleteFunction(buildingFunction)
  }

  async function handleConfirmDelete() {
    if (!sdk || !canEdit || !pendingDeleteFunction) {
      return
    }

    const buildingFunction = pendingDeleteFunction

    try {
      setBusyAction('delete')
      setNotice(null)
      setSettingsError(null)

      const response = await withTimeout(
        sdk.settings.buildingFunctions.delete({
          id: buildingFunction.id,
        }),
        'Forma.settings.buildingFunctions.delete()',
      )

      applySettingsResponse(response)

      if (editingId === buildingFunction.id) {
        clearEditing()
      }

      setPendingDeleteFunction(null)
      setNotice({
        severity: 'success',
        text: `Deleted "${buildingFunction.name}".`,
      })
    } catch (error) {
      const message = getErrorMessage(error)
      setPendingDeleteFunction(null)
      setSettingsError(message)
      setNotice({
        severity: 'error',
        text: `Could not delete the building function: ${message}`,
      })
    } finally {
      setBusyAction(null)
    }
  }

  return (
    <>
      <Box
        sx={{
          backgroundColor: 'background.default',
          color: 'text.primary',
          minHeight: '100vh',
        }}
      >
        <main className="app-shell">
          <Stack spacing={3}>
            <Stack spacing={1}>
              <Typography color="text.secondary" variant="overline">
                Autodesk Forma Site Design
              </Typography>
              <Typography component="h1" variant="h5">
                Building Functions and 2.5D Buildings
              </Typography>
              <Typography color="text.secondary" variant="body2">
                Manage project-level building functions and create persistent 2.5D
                buildings through Autodesk Forma.
              </Typography>
              <Typography color="text.secondary" variant="body2">
                {hostOrigin
                  ? `Project ${projectId ?? 'Unavailable'} · ${sdk ? (canEdit ? 'Edit access available' : 'Read-only access') : 'Connecting...'}`
                  : 'Host-only actions stay disabled until this page is opened inside Autodesk Forma.'}
              </Typography>
            </Stack>

            <Alert severity={hostBanner.severity}>
              <Typography variant="subtitle2">{hostBanner.title}</Typography>
              <Typography variant="body2">{hostBanner.text}</Typography>
            </Alert>

            {notice ? (
              <Alert onClose={() => setNotice(null)} severity={notice.severity}>
                {notice.text}
              </Alert>
            ) : null}

            <Box component="section" sx={sectionSx}>
              <Stack spacing={2}>
                <Stack
                  alignItems={{ sm: 'center', xs: 'flex-start' }}
                  direction={{ sm: 'row', xs: 'column' }}
                  justifyContent="space-between"
                  spacing={2}
                >
                  <Stack spacing={1}>
                    <Typography component="h2" variant="h6">
                      Current Building Functions
                    </Typography>
                    <Typography color="text.secondary" variant="body2">
                      {buildingFunctions.length} total
                      {hasCustomFunctions
                        ? ' · Includes custom functions'
                        : ' · Built-in functions only'}
                    </Typography>
                  </Stack>

                  <Button
                    disabled={!sdk}
                    loading={busyAction === 'refresh'}
                    onClick={() => {
                      void handleRefresh()
                    }}
                    variant="outlined"
                  >
                    Refresh
                  </Button>
                </Stack>

                <Typography color="text.secondary" variant="body2">
                  Built-in defaults are immutable. Custom functions can be edited or
                  removed when edit access is available.
                </Typography>

                {settingsError ? <Alert severity="error">{settingsError}</Alert> : null}

                {settingsLoading ? (
                  <Stack alignItems="center" spacing={2} sx={{ paddingY: 3 }}>
                    <CircularProgress />
                    <Typography color="text.secondary" variant="body2">
                      Loading building functions...
                    </Typography>
                  </Stack>
                ) : buildingFunctions.length === 0 ? (
                  <Alert severity="info">
                    No building functions were returned by Autodesk Forma.
                  </Alert>
                ) : (
                  <Stack spacing={0}>
                    {buildingFunctions.map((buildingFunction, index) => {
                      const builtIn = isBuiltInFunction(buildingFunction)
                      const selected = selectedBuildingFunction?.id === buildingFunction.id

                      return (
                        <Box
                          key={buildingFunction.id}
                          sx={{
                            borderColor: 'divider',
                            borderTop: index === 0 ? 0 : 1,
                            paddingTop: index === 0 ? 0 : 2,
                            marginTop: index === 0 ? 0 : 2,
                          }}
                        >
                          <Stack spacing={1}>
                            <Stack
                              alignItems={{ sm: 'center', xs: 'flex-start' }}
                              direction={{ sm: 'row', xs: 'column' }}
                              justifyContent="space-between"
                              spacing={2}
                            >
                              <Stack spacing={1}>
                                <Typography variant="subtitle1">
                                  {buildingFunction.name}
                                </Typography>
                                <Typography color="text.secondary" variant="body2">
                                  {buildingFunction.id} · {builtIn ? 'Built-in' : 'Custom'}
                                  {selected ? ' · Selected for editing' : ''}
                                </Typography>
                              </Stack>

                              {builtIn ? (
                                <Typography color="text.secondary" variant="body2">
                                  Immutable default
                                </Typography>
                              ) : (
                                <Stack direction="row" spacing={1}>
                                  <Button
                                    disabled={!canMutate}
                                    onClick={() => {
                                      handleSelectForUpdate(buildingFunction)
                                    }}
                                    variant={selected ? 'contained' : 'outlined'}
                                  >
                                    {selected ? 'Selected' : 'Edit'}
                                  </Button>
                                  <Button
                                    color="error"
                                    disabled={!canMutate}
                                    onClick={() => {
                                      handleRequestDelete(buildingFunction)
                                    }}
                                    variant="text"
                                  >
                                    Delete
                                  </Button>
                                </Stack>
                              )}
                            </Stack>

                            <Stack
                              alignItems="center"
                              direction="row"
                              spacing={1}
                            >
                              <Box
                                sx={{
                                  backgroundColor:
                                    buildingFunction.color ?? 'transparent',
                                  border: 1,
                                  borderColor: 'divider',
                                  borderRadius: '999px',
                                  flexShrink: 0,
                                  height: 16,
                                  width: 16,
                                }}
                              />
                              <Typography color="text.secondary" variant="body2">
                                {buildingFunction.color ?? 'No custom color'}
                              </Typography>
                            </Stack>
                          </Stack>
                        </Box>
                      )
                    })}
                  </Stack>
                )}
              </Stack>
            </Box>

            <Box component="section" sx={sectionSx}>
              <Stack spacing={2}>
                <Stack spacing={1}>
                  <Typography component="h2" variant="h6">
                    Create 2.5D Building
                  </Typography>
                  <Typography color="text.secondary" variant="body2">
                    Create a persistent floor-stack building and click a point in the
                    scene to place it in the current proposal.
                  </Typography>
                </Stack>

                {settingsLoading ? (
                  <Alert severity="info">
                    Loading building functions for 2.5D building creation...
                  </Alert>
                ) : buildingFunctions.length === 0 ? (
                  <Alert severity="info">
                    Load at least one building function before creating a 2.5D building.
                  </Alert>
                ) : null}

                {creationStatus ? <Alert severity="info">{creationStatus}</Alert> : null}

                {buildingFunctions.length > 0 ? (
                  <Box component="form" onSubmit={handleCreateBuilding}>
                    <Stack spacing={2}>
                      <TextField
                        disabled={!canCreateBuildings}
                        fullWidth
                        label="Building name"
                        onChange={(event) => {
                          setCreateBuildingName(event.target.value)
                        }}
                        placeholder="Mixed Use Test Tower"
                        required
                        value={createBuildingName}
                      />
                      <Stack spacing={1}>
                        <Typography color="text.secondary" variant="body2">
                          Building function
                        </Typography>
                        <Select
                          disabled={!canCreateBuildings}
                          fullWidth
                          onChange={(event) => {
                            setCreateFunctionId(String(event.target.value))
                          }}
                          value={createFunctionId}
                          variant="outlined"
                        >
                          {buildingFunctions.map((buildingFunction) => (
                            <MenuItem key={buildingFunction.id} value={buildingFunction.id}>
                              {buildingFunction.name} ({buildingFunction.id})
                            </MenuItem>
                          ))}
                        </Select>
                      </Stack>

                      <Stack
                        direction={{ md: 'row', xs: 'column' }}
                        spacing={2}
                      >
                        <TextField
                          disabled={!canCreateBuildings}
                          fullWidth
                          inputProps={{ min: 1, step: 0.1 }}
                          label="Width"
                          onChange={(event) => {
                            setCreateWidth(event.target.value)
                          }}
                          type="number"
                          value={createWidth}
                        />
                        <TextField
                          disabled={!canCreateBuildings}
                          fullWidth
                          inputProps={{ min: 1, step: 0.1 }}
                          label="Depth"
                          onChange={(event) => {
                            setCreateDepth(event.target.value)
                          }}
                          type="number"
                          value={createDepth}
                        />
                      </Stack>

                      <Stack
                        direction={{ md: 'row', xs: 'column' }}
                        spacing={2}
                      >
                        <TextField
                          disabled={!canCreateBuildings}
                          fullWidth
                          inputProps={{ min: 1, step: 1 }}
                          label="Floor count"
                          onChange={(event) => {
                            setCreateFloorCount(event.target.value)
                          }}
                          type="number"
                          value={createFloorCount}
                        />
                        <TextField
                          disabled={!canCreateBuildings}
                          fullWidth
                          inputProps={{ min: 0.1, step: 0.1 }}
                          label="Floor height"
                          onChange={(event) => {
                            setCreateFloorHeight(event.target.value)
                          }}
                          type="number"
                          value={createFloorHeight}
                        />
                      </Stack>

                      <Stack spacing={1}>
                        <Stack
                          alignItems="center"
                          direction="row"
                          spacing={1}
                          useFlexGap
                          flexWrap="wrap"
                        >
                          <Box
                            sx={{
                              backgroundColor:
                                selectedCreationFunction?.color ?? 'transparent',
                              border: 1,
                              borderColor: 'divider',
                              borderRadius: '999px',
                              flexShrink: 0,
                              height: 20,
                              width: 20,
                            }}
                          />
                          <Typography color="text.secondary" variant="body2">
                            {selectedCreationFunction
                              ? `${selectedCreationFunction.name} (${selectedCreationFunction.id})`
                              : 'Select a building function'}
                          </Typography>
                        </Stack>
                        <Typography color="text.secondary" variant="body2">
                          {selectedCreationFunction
                            ? isBuiltInFunction(selectedCreationFunction)
                              ? 'Built-in function'
                              : 'Custom function'
                            : ''}
                        </Typography>
                      </Stack>

                      <Box>
                        <Button
                          disabled={!canCreateBuildings}
                          loading={busyAction === 'create-building'}
                          type="submit"
                          variant="contained"
                        >
                          Create and place building
                        </Button>
                      </Box>
                    </Stack>
                  </Box>
                ) : null}
              </Stack>
            </Box>

            <Box component="section" sx={sectionSx}>
              <Stack spacing={2}>
                <Stack spacing={1}>
                  <Typography component="h2" variant="h6">
                    Add Custom Function
                  </Typography>
                  <Typography color="text.secondary" variant="body2">
                    Create a new project-level building function.
                  </Typography>
                </Stack>

                <Box component="form" onSubmit={handleAdd}>
                  <Stack spacing={2}>
                    <TextField
                      disabled={!canMutate}
                      fullWidth
                      label="Name"
                      onChange={(event) => {
                        setNewName(event.target.value)
                      }}
                      placeholder="Retail"
                      required
                      value={newName}
                    />
                    <TextField
                      disabled={!canMutate}
                      error={Boolean(newColor.trim()) && newColorPreview === null}
                      fullWidth
                      helperText={
                        newColor.trim() && newColorPreview === null
                          ? 'Use a 6-digit hex value like #FF5733.'
                          : 'Optional. Leave blank to use Forma defaults.'
                      }
                      label="Color"
                      onChange={(event) => {
                        setNewColor(event.target.value)
                      }}
                      placeholder="#FF5733"
                      value={newColor}
                    />
                    <Stack spacing={1}>
                      <Stack
                        alignItems="center"
                        direction="row"
                        spacing={1}
                      >
                        <Box
                          sx={{
                            backgroundColor: newColorPreview ?? 'transparent',
                            border: 1,
                            borderColor: 'divider',
                            borderRadius: '999px',
                            flexShrink: 0,
                            height: 20,
                            width: 20,
                          }}
                        />
                        <Typography color="text.secondary" variant="body2">
                          {newColorPreview ?? 'No custom color'}
                        </Typography>
                      </Stack>

                      <Box>
                        <Button
                          disabled={!canMutate}
                          loading={busyAction === 'add'}
                          type="submit"
                          variant="contained"
                        >
                          Add function
                        </Button>
                      </Box>
                    </Stack>
                  </Stack>
                </Box>
              </Stack>
            </Box>

            <Box component="section" sx={sectionSx}>
              <Stack spacing={2}>
                <Stack spacing={1}>
                  <Typography component="h2" variant="h6">
                    Update Selected Function
                  </Typography>
                  <Typography color="text.secondary" variant="body2">
                    Select a custom function from the list above, then update its name
                    or color.
                  </Typography>
                </Stack>

                {selectedBuildingFunction ? (
                  <Box component="form" onSubmit={handleUpdate}>
                    <Stack spacing={2}>
                      <TextField
                        disabled
                        fullWidth
                        label="Selected ID"
                        value={selectedBuildingFunction.id}
                      />
                      <TextField
                        disabled={!canMutate}
                        fullWidth
                        label="Name"
                        onChange={(event) => {
                          setEditName(event.target.value)
                        }}
                        required
                        value={editName}
                      />
                      <TextField
                        disabled={!canMutate}
                        error={Boolean(editColor.trim()) && editColorPreview === null}
                        fullWidth
                        helperText={
                          editColor.trim() && editColorPreview === null
                            ? 'Use a 6-digit hex value like #FF5733.'
                            : 'Optional. Leave blank to use Forma defaults.'
                        }
                        label="Color"
                        onChange={(event) => {
                          setEditColor(event.target.value)
                        }}
                        placeholder="#059669"
                        value={editColor}
                      />
                      <Stack spacing={1}>
                        <Stack
                          alignItems="center"
                          direction="row"
                          spacing={1}
                        >
                          <Box
                            sx={{
                              backgroundColor: editColorPreview ?? 'transparent',
                              border: 1,
                              borderColor: 'divider',
                              borderRadius: '999px',
                              flexShrink: 0,
                              height: 20,
                              width: 20,
                            }}
                          />
                          <Typography color="text.secondary" variant="body2">
                            {editColorPreview ?? 'No custom color'}
                          </Typography>
                        </Stack>

                        <Stack direction="row" spacing={1}>
                          <Button
                            disabled={isBusy}
                            onClick={clearEditing}
                            variant="outlined"
                          >
                            Clear
                          </Button>
                          <Button
                            disabled={!canMutate}
                            loading={busyAction === 'update'}
                            type="submit"
                            variant="contained"
                          >
                            Update function
                          </Button>
                        </Stack>
                      </Stack>
                    </Stack>
                  </Box>
                ) : (
                  <Alert severity="info">
                    Select a custom building function to edit it.
                  </Alert>
                )}
              </Stack>
            </Box>
          </Stack>
        </main>
      </Box>

      <Modal
        actions={
          <Stack direction="row" justifyContent="flex-end" spacing={1}>
            <Button
              disabled={busyAction === 'delete'}
              onClick={() => {
                setPendingDeleteFunction(null)
              }}
              variant="outlined"
            >
              Cancel
            </Button>
            <Button
              color="error"
              loading={busyAction === 'delete'}
              onClick={() => {
                void handleConfirmDelete()
              }}
              variant="contained"
            >
              Delete
            </Button>
          </Stack>
        }
        header="Delete custom building function"
        onClose={() => {
          if (busyAction !== 'delete') {
            setPendingDeleteFunction(null)
          }
        }}
        open={pendingDeleteFunction !== null}
        subheader={pendingDeleteFunction?.name}
      >
        <Box>
          <Stack spacing={1}>
            <Typography variant="body2">
              This removes the custom building function from the current project.
            </Typography>
            {pendingDeleteFunction ? (
              <Typography color="text.secondary" variant="body2">
                ID: {pendingDeleteFunction.id}
              </Typography>
            ) : null}
          </Stack>
        </Box>
      </Modal>
    </>
  )
}

export default App
