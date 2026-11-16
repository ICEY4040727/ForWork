import type { DesktopFileActions } from '@forwork/client-ui-file-actions/types'
/** Typed preload operations exposed only by the Electron shell. */

export type { DesktopChannelState, DesktopChannelChangeResult } from '@forwork/client-ui-settings-general/types'
import type { DesktopChannelBridge, DesktopChannelState } from '@forwork/client-ui-settings-general/types'
import type { DesktopKeyboardApi, DesktopShortcutsApi } from '@forwork/client-shortcuts/protocol'
import type { IpcMainInvokeEvent } from 'electron'
import type { DesktopBrowserBridge } from '@forwork/client-ui-sidebar-browser/types'
import type { DesktopFileDrag } from '@forwork/util-file-drag'

/** IPC channel names kept private to the desktop application bundle. */
export const DESKTOP_IPC = {
  fileSave: 'forwork-desktop:file-save',
  clipboardImageRead: 'forwork-desktop:clipboard-image-read',
  shortcutsInput: 'forwork-desktop:shortcuts-input',
  shortcutsCloseWindow: 'forwork-desktop:shortcuts-close-window',
  shortcutsGet: 'forwork-desktop:shortcuts-get',
  shortcutsEdit: 'forwork-desktop:shortcuts-edit',
  shortcutsChanged: 'forwork-desktop:shortcuts-changed',
  shortcutsRecording: 'forwork-desktop:shortcuts-recording',
  boot: 'forwork-desktop:boot',
  enterWorkspace: 'forwork-desktop:enter-workspace',
  onboardingActive: 'forwork-desktop:onboarding-active',
  onboardingApiKey: 'forwork-desktop:onboarding-api-key',
  bootFailed: 'forwork-desktop:boot-failed',
  browserAcquire: 'forwork-desktop:browser-acquire',
  browserRelease: 'forwork-desktop:browser-release',
  browserOpenRequested: 'forwork-desktop:browser-open-requested',
  directoryPick: 'forwork-desktop:directory-pick',
  fileDragPrepare: 'forwork-desktop:file-drag-prepare',
  fileDragCancel: 'forwork-desktop:file-drag-cancel',
  fileDragStart: 'forwork-desktop:file-drag-start',
  fileDragRelease: 'forwork-desktop:file-drag-release',
  deviceInfo: 'forwork-desktop:device-info',
  localeBootstrap: 'forwork-desktop:locale-bootstrap',
  localeChanged: 'forwork-desktop:locale-changed',
  updatesChannelGet: 'forwork-desktop:updates-channel-get',
  updatesChannelSet: 'forwork-desktop:updates-channel-set',
  updatesChannelChanged: 'forwork-desktop:updates-channel-changed',
  updatesStatus: 'forwork-desktop:updates-status',
  updatesOpen: 'forwork-desktop:updates-open',
  updatesPresentation: 'forwork-desktop:updates-presentation',
  nativeThemeSet: 'forwork-desktop:native-theme-set',
  windowFullscreen: 'forwork-desktop:window-fullscreen',
  windowsAppearance: 'forwork-desktop:windows-appearance',
  windowsMenu: 'forwork-desktop:windows-menu',
} as const

/** Desktop release update state rendered by desktop-owned UI. */
export type DesktopUpdatePreparationFailureKind = 'stop-failed' | 'tasks-changed' | 'tasks-unavailable' | 'fullscreen-timeout'

export interface DesktopUpdateState {
  readonly phase: 'idle' | 'checking' | 'available' | 'downloading' | 'verifying' | 'installing' | 'ready' | 'error'
  readonly version?: string
  readonly message?: string
  /** Main-owned diagnostics without subprocess output or credentials; hidden until expanded. */
  readonly technicalDetails?: string
  readonly percent?: number
  readonly failedOperation?: 'check' | 'download' | 'install' | 'configuration'
  /** Main-owned preparation cause; UI wording is selected by the active locale. */
  readonly preparationFailure?: DesktopUpdatePreparationFailureKind
  /** Current filesystem failure, independent of previous cache cleanup attempts. */
  readonly downloadFailure?: 'space' | 'permission' | 'busy'
}

/** Classified failure copy selected by the Web locale without exposing raw updater diagnostics. */
export type DesktopUpdateFailureKind =
  | 'configuration'
  | 'check'
  | 'check-network'
  | 'download'
  | 'download-network'
  | 'download-space'
  | 'download-permission'
  | 'download-busy'
  | 'install'
  | 'install-network'
  | 'stop-failed'
  | 'tasks-changed'
  | 'tasks-unavailable'
  | 'fullscreen-timeout'

/** Semantic status content; actions open main-process confirmation dialogs only. */
export interface DesktopUpdatePresentation {
  readonly phase: DesktopUpdateState['phase']
  readonly version?: string
  readonly percent?: number
  readonly failure?: DesktopUpdateFailureKind
  readonly configurationError?: DesktopChannelState['error']
}

/** Product documents cannot supply update versions, package URLs, or installation authorization. */
export interface DshDesktopProductApi {
  readonly fileActions: DesktopFileActions
  readonly protocolVersion: 1
  readonly browser: DesktopBrowserBridge
  readonly fileDrag: DesktopFileDrag
  readonly keyboard: DesktopKeyboardApi
  readonly shortcuts: DesktopShortcutsApi
  /**
   * Local machine description for the feedback questionnaire.
   * @returns `name=value` fields separated by `; `, with no hostname, user name, or serial number.
   */
  deviceInfo(): Promise<string>
  readonly updates: {
    readonly channels: DesktopChannelBridge
    status(): Promise<DesktopUpdatePresentation>
    open(): Promise<void>
    subscribe(listener: (state: DesktopUpdatePresentation) => void): () => void
  }
}

/** Scheme of Desktop-owned application documents. */
export const SCHEME = 'dsh-app'

/**
 * Reject IPC outside the allowed Desktop document origins.
 * @param event - IPC caller whose frame URL supplies the origin.
 * @param hostnames - Desktop document hosts allowed for this operation.
 */
export function assertDesktopSender(event: IpcMainInvokeEvent, hostnames: readonly string[]): void {
  const senderFrame = event.senderFrame
  if (senderFrame === null) throw new Error('forwork desktop: rejected IPC without a sender frame')
  const url = new URL(senderFrame.url)
  if (url.protocol !== `${SCHEME}:` || !hostnames.includes(url.hostname)) {
    throw new Error('forwork desktop: rejected IPC from an unowned renderer')
  }
}
