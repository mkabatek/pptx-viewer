import { C as CollaborationConfig, d as ConnectionStatus, U as UserPresence, e as CollaborationContextValue } from '../PowerPointViewer-C5jGuKGB.mjs';
export { A as AccessibilityIssue, f as AnimationPresetOption, h as CanvasSize, i as CollaborationRole, j as ConnectorArrowOption, k as ConnectorGeometryOption, l as ConnectorGeometryType, m as ConnectorPathGeometry, D as DragState, n as DrawingTool, E as EditorHistorySnapshot, o as ElementBounds, p as ElementClipboardPayload, q as ElementContextMenuAction, r as ElementContextMenuState, F as FileViewerHandle, M as MarqueeSelectionState, s as ParsedTableCell, t as ParsedTableData, P as PowerPointViewer, b as PowerPointViewerHandle, c as PowerPointViewerProps, u as PresentationAnimationRuntime, R as ResizeHandle, v as ResizeState, S as ShapeAdjustmentDragState, w as ShapeAdjustmentHandleDescriptor, x as ShapePreset, y as ShapeQuickStyle, z as ShortcutReferenceItem, B as SlideAlignment, G as SlideSectionGroup, H as SlideTransitionOption, I as StrokeDashOption, J as SupportedShapeType, T as TableCellEditorState, K as ToolbarSection, L as ViewerMode, g as getAnimationInitialStyle } from '../PowerPointViewer-C5jGuKGB.mjs';
import { PptxThemePreset, PptxThemeColorScheme, PptxThemeFontScheme, PptxHandler, PptxData } from 'pptx-viewer-core';
import React__default, { RefObject } from 'react';
import { Awareness } from 'y-protocols/awareness';
import { Doc } from 'yjs';

/** Returns true if the current page was opened as an audience tab. */
declare function isAudienceTab(): boolean;

/**
 * audience-content-store — IndexedDB-based storage for sharing PPTX content
 * between the presenter tab and audience tab.
 *
 * When the presenter opens an audience window, the PPTX bytes are stored in
 * IndexedDB. The audience tab retrieves them on load and then cleans up.
 */
/**
 * Store PPTX content bytes so the audience tab can retrieve them.
 * Called by the presenter before opening the audience window.
 */
declare function storeAudienceContent(content: ArrayBuffer | Uint8Array): Promise<void>;
/**
 * Load PPTX content bytes stored by the presenter tab.
 * Returns `null` if nothing is stored.
 */
declare function loadAudienceContent(): Promise<Uint8Array | null>;
/**
 * Remove stored audience content (cleanup).
 */
declare function clearAudienceContent(): Promise<void>;

interface UseThemeSwitchingInput {
    /** Ref to the PptxHandler instance. */
    handlerRef: RefObject<PptxHandler | null>;
    /** Current parsed presentation data (null when nothing is loaded). */
    data: PptxData | null;
    /** Callback to update the presentation data after theme switch. */
    onDataChange: (newData: PptxData) => void;
    /** Optional callback fired when theme switch completes successfully. */
    onThemeChanged?: (preset: PptxThemePreset) => void;
}
interface ThemeSwitchingResult {
    /** All available built-in theme presets. */
    presets: readonly PptxThemePreset[];
    /**
     * Apply a theme preset to the current presentation.
     * Updates both the in-memory ZIP and all resolved element colours.
     */
    switchToPreset: (preset: PptxThemePreset) => Promise<void>;
    /**
     * Apply a custom colour scheme (and optional font scheme) to the
     * current presentation.
     */
    switchToCustom: (colorScheme: PptxThemeColorScheme, fontScheme?: PptxThemeFontScheme, themeName?: string) => Promise<void>;
    /**
     * Get the preset matching the current presentation theme (if any).
     * Returns undefined if the current theme does not match a built-in preset.
     */
    currentPreset: PptxThemePreset | undefined;
}
/**
 * React hook providing theme switching capabilities for the PowerPoint viewer.
 *
 * @example
 * ```tsx
 * const { presets, switchToPreset, currentPreset } = useThemeSwitching({
 *   handlerRef,
 *   data,
 *   onDataChange: setData,
 * });
 *
 * return (
 *   <div>
 *     {presets.map(preset => (
 *       <button
 *         key={preset.id}
 *         onClick={() => switchToPreset(preset)}
 *         aria-pressed={preset.id === currentPreset?.id}
 *       >
 *         {preset.name}
 *       </button>
 *     ))}
 *   </div>
 * );
 * ```
 */
declare function useThemeSwitching(input: UseThemeSwitchingInput): ThemeSwitchingResult;

interface UseYjsProviderInput {
    config: CollaborationConfig;
}
interface UseYjsProviderResult {
    /** Current WebSocket connection status. */
    status: ConnectionStatus;
    /** The Yjs awareness instance (null until connected). */
    awareness: Awareness | null;
    /** The Yjs document (null until initialised). */
    doc: Doc | null;
    /** Local awareness client ID. */
    clientId: number | null;
    /** Manually retry the connection after a timeout or error. */
    retry: () => void;
}
/**
 * Lazily loads `yjs` and `y-websocket`, creates a Y.Doc and
 * WebSocketProvider, and tracks the connection lifecycle.
 *
 * If the connection does not succeed within {@link CONNECTION_TIMEOUT_MS},
 * the provider is torn down and status moves to `'error'`. The consumer
 * can call `retry()` to attempt a fresh connection.
 *
 * The Yjs packages are dynamically imported so they are fully
 * tree-shaken when collaboration is not enabled.
 */
declare function useYjsProvider({ config }: UseYjsProviderInput): UseYjsProviderResult;

interface UsePresenceTrackingInput {
    awareness: Awareness | null;
    localClientId: number | null;
    userName: string;
    userColor: string;
    userAvatar?: string;
    /** Role in the session (broadcaster, viewer, or collaborator). */
    role?: string;
    canvasWidth: number;
    canvasHeight: number;
}
interface UsePresenceTrackingResult {
    /** Presence data for all remote users (excludes local). */
    remoteUsers: UserPresence[];
    /** Broadcast a partial presence update for the local user. */
    broadcastPresence: (update: Partial<Omit<UserPresence, 'clientId'>>) => void;
}
declare function usePresenceTracking({ awareness, localClientId, userName, userColor, userAvatar, role, canvasWidth, canvasHeight, }: UsePresenceTrackingInput): UsePresenceTrackingResult;

/**
 * useCollaborativeState — Composes the Yjs provider and presence tracking
 * into a single hook for the collaboration system.
 *
 * This is the primary hook consumed by the `CollaborationProvider` context.
 * It orchestrates:
 * - Yjs WebSocket connection lifecycle
 * - Presence tracking (broadcast + receive)
 * - Connection status
 *
 * @module collaboration/useCollaborativeState
 */

interface UseCollaborativeStateInput {
    config: CollaborationConfig;
    canvasWidth: number;
    canvasHeight: number;
}
declare function useCollaborativeState({ config, canvasWidth, canvasHeight, }: UseCollaborativeStateInput): CollaborationContextValue;

interface UseCollaborativeHistoryInput {
    /** The local user's client ID (from Yjs awareness). */
    localClientId: number | null;
    /** Standard history undo function. */
    handleUndo: () => void;
    /** Standard history redo function. */
    handleRedo: () => void;
    /** Whether undo is available. */
    canUndo: boolean;
    /** Whether redo is available. */
    canRedo: boolean;
}
interface UseCollaborativeHistoryResult {
    /** Undo the last local change. */
    handleUndo: () => void;
    /** Redo the last undone local change. */
    handleRedo: () => void;
    /** Whether undo is available. */
    canUndo: boolean;
    /** Whether redo is available. */
    canRedo: boolean;
}
declare function useCollaborativeHistory({ localClientId: _localClientId, handleUndo, handleRedo, canUndo, canRedo, }: UseCollaborativeHistoryInput): UseCollaborativeHistoryResult;

/**
 * CollaborationProvider — React context provider for real-time collaboration.
 *
 * Wraps the viewer content and exposes collaboration state (connection status,
 * remote user presence, broadcast function) to all child components via
 * `useCollaboration()`.
 *
 * This component is only rendered when `collaboration` config is provided to
 * `PowerPointViewer`, ensuring zero bundle/runtime cost when unused.
 *
 * @module collaboration/CollaborationProvider
 */

interface CollaborationProviderProps {
    config: CollaborationConfig;
    canvasWidth: number;
    canvasHeight: number;
    children: React__default.ReactNode;
}
declare function CollaborationProvider({ config, canvasWidth, canvasHeight, children, }: CollaborationProviderProps): React__default.ReactElement;

/**
 * RemoteUserCursors — Renders other users' cursors as an SVG overlay
 * on the slide canvas.
 *
 * Each remote user's cursor is drawn as a coloured pointer arrow with
 * their username label. Only cursors on the same slide as the local
 * user are displayed.
 *
 * @module collaboration/RemoteUserCursors
 */

interface RemoteUserCursorsProps {
    /** Presence data for remote users. */
    remoteUsers: UserPresence[];
    /** The slide index the local user is currently viewing. */
    activeSlideIndex: number;
    /** Canvas width in CSS px (for SVG viewBox). */
    canvasWidth: number;
    /** Canvas height in CSS px (for SVG viewBox). */
    canvasHeight: number;
}
declare function RemoteUserCursors({ remoteUsers, activeSlideIndex, canvasWidth, canvasHeight, }: RemoteUserCursorsProps): React__default.ReactElement | null;

/**
 * UserAvatarBar — Displays connected collaborators as a row of avatar circles
 * in the toolbar area.
 *
 * Shows up to 5 avatar circles with a "+N" overflow indicator.
 * Each circle shows the user's avatar image (if available) or their initials.
 *
 * @module collaboration/UserAvatarBar
 */

interface UserAvatarBarProps {
    /** Remote user presence data. */
    remoteUsers: UserPresence[];
    /** Local user display name. */
    localUserName: string;
    /** Local user's colour. */
    localUserColor: string;
    /** Local user's avatar URL. */
    localUserAvatar?: string;
    /** Connection status. */
    status: ConnectionStatus;
    /** Maximum visible avatars before showing overflow (default: 5). */
    maxVisible?: number;
}
declare function UserAvatarBar({ remoteUsers, localUserName, localUserColor, localUserAvatar, status, maxVisible, }: UserAvatarBarProps): React__default.ReactElement | null;

/**
 * CollaborationStatusIndicator — A small status pill that shows the
 * WebSocket connection state and connected user count.
 *
 * Designed to sit in the status bar area at the bottom of the viewer.
 *
 * @module collaboration/CollaborationStatusIndicator
 */

interface CollaborationStatusIndicatorProps {
    /** Current WebSocket connection status. */
    status: ConnectionStatus;
    /** Number of connected users (including local). */
    connectedCount: number;
    /** Callback to retry the connection (shown for error state). */
    onRetry?: () => void;
}
declare function CollaborationStatusIndicator({ status, connectedCount, onRetry, }: CollaborationStatusIndicatorProps): React__default.ReactElement;

export { CollaborationConfig, CollaborationContextValue, CollaborationProvider, CollaborationStatusIndicator, type CollaborationStatusIndicatorProps, ConnectionStatus, RemoteUserCursors, type RemoteUserCursorsProps, type ThemeSwitchingResult, type UseCollaborativeHistoryResult, type UseCollaborativeStateInput, type UsePresenceTrackingResult, type UseThemeSwitchingInput, UserAvatarBar, type UserAvatarBarProps, UserPresence, clearAudienceContent, isAudienceTab, loadAudienceContent, storeAudienceContent, useCollaborativeHistory, useCollaborativeState, usePresenceTracking, useThemeSwitching, useYjsProvider };
