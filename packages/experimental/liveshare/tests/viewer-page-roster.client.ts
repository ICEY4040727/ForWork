/**
 * Test-only client roster of the assembled viewer page.
 *
 * A served viewer page is the owner's own Web shell with the owner's client plugins; this roster is the
 * subset the viewer page specs need: the wire plugin, the Gateway, remote access, the Session/Workspace controllers,
 * the menu service, presentation plugins, and the Liveshare `./client` face, which mounts the viewer surfaces when the
 * carrier names a share. Every row is the package's own `/client` face registered as a module-table
 * factory. It must be imported only after {@link prepareViewerRemoteFaces} answered the generated
 * `/remote` faces these plugins import.
 *
 * @module @forwork/experimental-liveshare/tests/viewer-page-roster
 */

import * as apiGateway from '@forwork/api-gateway/client'
import * as apiRemoteAccess from '@forwork/api-remote-access/client'
import * as apiRemotes from '@forwork/api-remotes/client'
import * as apiSessionController from '@forwork/api-session-controller/client'
import * as apiWorkspaceController from '@forwork/api-workspace-controller/client'
import * as clientConnection from '@forwork/client-connection/client'
import * as clientFileUpload from '@forwork/client-file-upload/client'
import * as clientLocale from '@forwork/client-locale/client'
import * as clientModules from '@forwork/client-modules/client'
import * as clientResources from '@forwork/client-resources/client'
import * as clientShortcuts from '@forwork/client-shortcuts/client'
import * as uiAttachment from '@forwork/client-ui-attachment/client'
import * as uiChat from '@forwork/client-ui-chat/client'
import * as uiContextMenu from '@forwork/client-ui-context-menu/client'
import * as uiConversation from '@forwork/client-ui-conversation/client'
import * as uiDeliverables from '@forwork/client-ui-deliverables/client'
import * as uiFileDrag from '@forwork/client-ui-file-drag/client'
import * as uiLayout from '@forwork/client-ui-layout/client'
import * as uiRenderer from '@forwork/client-ui-renderer/client'
import * as uiSession from '@forwork/client-ui-session/client'
import * as uiSettings from '@forwork/client-ui-settings/client'
import * as uiSidebar from '@forwork/client-ui-sidebar/client'
import * as uiSidebarRight from '@forwork/client-ui-sidebar-right/client'
import * as uiTheme from '@forwork/client-ui-theme/client'
import * as uiTool from '@forwork/client-ui-tool/client'
import * as uiTrajectory from '@forwork/client-ui-trajectory/client'
import * as uiWorkspace from '@forwork/client-ui-workspace/client'
import * as typertRegistry from '@forwork/typert-registry/client'
import * as liveshare from '../src/client/index.ts'

/** One module-table row: the entry id the Loader creates, and that plugin's client exports. */
export interface ViewerPageRosterRow {
  readonly id: string
  readonly exports: Readonly<Record<string, unknown>>
}

/** Entry id of the module-system bootstrap whose exports the facade hands to the module system. */
export const VIEWER_PAGE_BOOTSTRAP_ID = '@forwork/client-modules'

/** The viewer page roster; Cordis composes these entries by service injection, so order is irrelevant. */
export const VIEWER_PAGE_ROSTER: readonly ViewerPageRosterRow[] = [
  { id: '@forwork/typert-registry', exports: typertRegistry },
  { id: '@forwork/experimental-liveshare', exports: liveshare },
  { id: VIEWER_PAGE_BOOTSTRAP_ID, exports: clientModules },
  { id: '@forwork/client-connection', exports: clientConnection },
  { id: '@forwork/api-gateway', exports: apiGateway },
  { id: '@forwork/api-remote-access', exports: apiRemoteAccess },
  { id: '@forwork/api-remotes', exports: apiRemotes },
  { id: '@forwork/api-session-controller', exports: apiSessionController },
  { id: '@forwork/api-workspace-controller', exports: apiWorkspaceController },
  { id: '@forwork/client-file-upload', exports: clientFileUpload },
  { id: '@forwork/client-locale', exports: clientLocale },
  { id: '@forwork/client-resources', exports: clientResources },
  { id: '@forwork/client-shortcuts', exports: clientShortcuts },
  { id: '@forwork/client-ui-renderer', exports: uiRenderer },
  { id: '@forwork/client-ui-theme', exports: uiTheme },
  { id: '@forwork/client-ui-settings', exports: uiSettings },
  { id: '@forwork/client-ui-layout', exports: uiLayout },
  { id: '@forwork/client-ui-session', exports: uiSession },
  { id: '@forwork/client-ui-sidebar', exports: uiSidebar },
  { id: '@forwork/client-ui-conversation', exports: uiConversation },
  { id: '@forwork/client-ui-chat', exports: uiChat },
  { id: '@forwork/client-ui-context-menu', exports: uiContextMenu },
  { id: '@forwork/client-ui-tool', exports: uiTool },
  { id: '@forwork/client-ui-attachment', exports: uiAttachment },
  { id: '@forwork/client-ui-file-drag', exports: uiFileDrag },
  { id: '@forwork/client-ui-sidebar-right', exports: uiSidebarRight },
  { id: '@forwork/client-ui-trajectory', exports: uiTrajectory },
  { id: '@forwork/client-ui-deliverables', exports: uiDeliverables },
  { id: '@forwork/client-ui-workspace', exports: uiWorkspace },
]
