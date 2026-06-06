import { PptxSlide, PptxElement, ShapeStyle, PptxAnimationPreset, ConnectorArrowType, XmlObject, PptxElementAnimation, PptxSlideTransition, StrokeDashType, PptxNativeAnimation } from 'pptx-viewer-core';
import * as React from 'react';
import React__default from 'react';
import { Doc } from 'yjs';

/**
 * Core data-model types for the PowerPoint viewer/editor plugin.
 *
 * Framework-agnostic types (PptxElement, PptxSlide, TextStyle, etc.)
 * live in `pptx-viewer-core`. This module defines additional types used
 * by the React viewer layer for canvas layout, drag/resize interactions,
 * clipboard operations, undo/redo history, geometry, and shape quick styles.
 */

/**
 * The current operating mode of the PowerPoint viewer.
 *
 * - `"preview"` -- Read-only viewing with no editing controls.
 * - `"edit"` -- Full editing mode with toolbar, inspector, and canvas interaction.
 * - `"present"` -- Fullscreen presentation/slideshow mode with animations.
 * - `"master"` -- Slide master editing mode for template/layout manipulation.
 */
type ViewerMode = 'preview' | 'edit' | 'present' | 'master';
/**
 * Union of all shape preset types that the viewer can insert or render.
 * Maps to OOXML `a:prstGeom` preset values (e.g. "rect", "ellipse").
 * Includes basic shapes, arrows, connectors, and freeform paths.
 */
type SupportedShapeType = 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'rtTriangle' | 'diamond' | 'cylinder' | 'parallelogram' | 'trapezoid' | 'pentagon' | 'hexagon' | 'octagon' | 'chevron' | 'star5' | 'star6' | 'star8' | 'plus' | 'heart' | 'cloud' | 'sun' | 'moon' | 'pie' | 'plaque' | 'teardrop' | 'line' | 'rtArrow' | 'leftArrow' | 'upArrow' | 'downArrow' | 'connector' | 'freeform';
/**
 * Connector geometry presets corresponding to OOXML `a:prstGeom` connector types.
 * Determines the routing style (straight, bent with N segments, or curved).
 */
type ConnectorGeometryType = 'straightConnector1' | 'bentConnector2' | 'bentConnector3' | 'bentConnector4' | 'bentConnector5' | 'curvedConnector2' | 'curvedConnector3' | 'curvedConnector4' | 'curvedConnector5';
/** The pixel dimensions of the slide canvas (presentation slide size). */
interface CanvasSize {
    /** Width of the canvas in CSS pixels. */
    width: number;
    /** Height of the canvas in CSS pixels. */
    height: number;
}
/**
 * Tracks the state of an ongoing element drag operation.
 * Created on pointer-down and destroyed on pointer-up.
 * During the drag, position changes are applied directly to DOM elements
 * for performance (bypassing React state) and committed to state on completion.
 */
interface DragState {
    /** ID of the primary element being dragged. */
    elementId: string;
    /** Client X coordinate at the start of the drag. */
    startClientX: number;
    /** Client Y coordinate at the start of the drag. */
    startClientY: number;
    /** Original positions of all selected elements, keyed by element ID. */
    startPositionsById: Record<string, {
        x: number;
        y: number;
    }>;
    /** DOM elements cached at drag-start to avoid per-frame querySelector calls. */
    domEls: Map<string, HTMLElement>;
    /** Whether the pointer has actually moved since the drag started. */
    moved: boolean;
    /** Accumulated delta — stored during DOM-only drag, committed on pointerup. */
    lastDx: number;
    lastDy: number;
}
/** Handle used as the resize anchor point (corners + edge midpoints). */
type ResizeHandle = 'nw' | 'ne' | 'sw' | 'se' | 'n' | 's' | 'e' | 'w';
/**
 * Tracks the state of an ongoing element resize operation.
 * Uses direct DOM manipulation during the resize for performance,
 * then commits the final geometry to React state on pointer-up.
 */
interface ResizeState {
    /** ID of the element being resized. */
    elementId: string;
    /** Client X coordinate at the start of the resize. */
    startClientX: number;
    /** Client Y coordinate at the start of the resize. */
    startClientY: number;
    /** Original X position of the element. */
    startX: number;
    /** Original Y position of the element. */
    startY: number;
    /** Original width of the element. */
    startWidth: number;
    /** Original height of the element. */
    startHeight: number;
    /** Which handle initiated the resize. */
    handle: ResizeHandle;
    /** Whether the pointer has actually moved since the resize started. */
    moved: boolean;
    /** DOM element cached at resize-start to avoid per-frame querySelector calls. */
    domEl: HTMLElement | null;
    /** Accumulated final geometry — committed on pointerup. */
    lastX: number;
    lastY: number;
    lastWidth: number;
    lastHeight: number;
}
/**
 * Describes the position and value of a shape adjustment handle (yellow diamond).
 * Adjustment handles allow users to modify shape parameters like corner radius,
 * arrow width, etc. without a full resize.
 */
interface ShapeAdjustmentHandleDescriptor {
    key: string;
    left: number;
    top: number;
    value: number;
    cursor: string;
}
/**
 * Tracks the state of an ongoing shape adjustment handle drag.
 * Used when the user drags the yellow diamond handle to alter
 * shape-specific parameters (e.g. corner radius on a rounded rectangle).
 */
interface ShapeAdjustmentDragState {
    elementId: string;
    key: string;
    shapeType: string;
    startClientX: number;
    startClientY: number;
    startAdjustment: number;
    startWidth: number;
    startHeight: number;
    moved: boolean;
}
/** Payload stored when an element is copied/cut to the internal clipboard. */
interface ElementClipboardPayload {
    element: PptxElement;
    isTemplate: boolean;
}
/**
 * A snapshot of editor state captured for undo/redo history.
 * Contains the full slide deck state at a point in time so that
 * any editing operation can be reverted or replayed.
 */
interface EditorHistorySnapshot {
    width: number;
    height: number;
    activeSlideIndex: number;
    slides: PptxSlide[];
    templateElementsBySlideId: Record<string, PptxElement[]>;
    actionLabel?: string;
}
/** Axis-aligned bounding box for an element, used for hit-testing and alignment. */
interface ElementBounds {
    minX: number;
    minY: number;
    maxX: number;
    maxY: number;
}
/**
 * Geometry data for rendering a connector line as an SVG path.
 * Includes start/end coordinates, the SVG path data string,
 * and optional marker (arrowhead) identifiers.
 */
interface ConnectorPathGeometry {
    startX: number;
    startY: number;
    endX: number;
    endY: number;
    pathData: string;
    d?: string;
    viewBox?: string;
    startMarkerId?: string;
    endMarkerId?: string;
}
/** A named preset shape style used in the Quick Styles gallery. */
interface ShapeQuickStyle {
    name: string;
    style: Partial<ShapeStyle>;
}

/**
 * Theme configuration types for the PowerPoint viewer.
 *
 * All color values accept any valid CSS color string:
 * hex (`#6366f1`), rgb (`rgb(99 102 241)`), hsl (`hsl(239 84% 67%)`),
 * oklch (`oklch(0.585 0.233 277)`), named colors, etc.
 */
/**
 * Semantic color tokens for the viewer UI.
 *
 * These map to CSS custom properties (`--pptx-<token>`) and drive all
 * UI component colors. The naming follows the shadcn/ui convention so
 * that Tailwind + shadcn users get a familiar experience.
 */
interface ViewerThemeColors {
    /** Page / root background */
    background: string;
    /** Default text color */
    foreground: string;
    /** Card / panel surface */
    card: string;
    /** Text on card surfaces */
    cardForeground: string;
    /** Popover / dropdown surface */
    popover: string;
    /** Text inside popovers */
    popoverForeground: string;
    /** Primary action color (buttons, active indicators) */
    primary: string;
    /** Text on primary-colored backgrounds */
    primaryForeground: string;
    /** Secondary / subdued action color */
    secondary: string;
    /** Text on secondary backgrounds */
    secondaryForeground: string;
    /** Muted / disabled surface */
    muted: string;
    /** Text on muted surfaces (also used for secondary text) */
    mutedForeground: string;
    /** Accent / hover-highlight surface */
    accent: string;
    /** Text on accent surfaces */
    accentForeground: string;
    /** Destructive / danger action color */
    destructive: string;
    /** Text on destructive backgrounds */
    destructiveForeground: string;
    /** Default border color */
    border: string;
    /** Input field border color */
    input: string;
    /** Focus ring color */
    ring: string;
}
/**
 * Full viewer theme configuration.
 *
 * Every property is optional — unset values fall back to the built-in
 * dark theme defaults.
 */
interface ViewerTheme {
    /** Semantic UI colors. Each key maps to a `--pptx-<key>` CSS custom property. */
    colors?: Partial<ViewerThemeColors>;
    /** Base border-radius value (e.g. `"0.5rem"`, `"8px"`). */
    radius?: string;
    /**
     * Escape hatch: arbitrary CSS custom properties to set on the viewer
     * root element. Keys should include the `--` prefix.
     *
     * @example
     * ```ts
     * { "--my-custom-shadow": "0 4px 12px rgba(0,0,0,0.5)" }
     * ```
     */
    cssVars?: Record<string, string>;
}

/**
 * Collaboration types — Shared type definitions for the real-time
 * collaboration infrastructure (Yjs-backed CRDT sync, presence tracking,
 * collaborative editing).
 *
 * @module collaboration/types
 */

/**
 * Configuration for enabling real-time collaboration on a presentation.
 *
 * When provided to `PowerPointViewer`, the viewer wraps its content in a
 * `CollaborationProvider` and wires up presence tracking, remote cursors,
 * and CRDT-based state synchronisation.
 */
/** Role of a user within a collaboration or broadcast session. */
type CollaborationRole = 'collaborator' | 'broadcaster' | 'viewer';
interface CollaborationConfig {
    /** Unique identifier for the collaboration room (alphanumeric, hyphens, underscores). */
    roomId: string;
    /** WebSocket server URL for the Yjs provider (e.g. "wss://collab.example.com"). */
    serverUrl: string;
    /** Display name for the local user. */
    userName: string;
    /** Avatar URL for the local user (optional). */
    userAvatar?: string;
    /** Hex colour for the local user's cursor/presence indicator. */
    userColor?: string;
    /** Optional authentication token sent with the WebSocket handshake. */
    authToken?: string;
    /** Role in the session — defaults to `'collaborator'`. */
    role?: CollaborationRole;
}
/** Connection lifecycle states for the Yjs WebSocket provider. */
type ConnectionStatus = 'disconnected' | 'connecting' | 'connected' | 'error';
/**
 * Presence data broadcast to other participants.
 * Cursor position is relative to the slide canvas (0..canvasWidth, 0..canvasHeight).
 */
interface UserPresence {
    /** Unique client ID (assigned by Yjs awareness). */
    clientId: number;
    /** Sanitised display name. */
    userName: string;
    /** Optional avatar URL (validated). */
    userAvatar?: string;
    /** Hex colour for the user's cursor ring. */
    userColor: string;
    /** Slide index the user is currently viewing. */
    activeSlideIndex: number;
    /** Cursor X on the canvas (clamped to slide bounds). */
    cursorX: number;
    /** Cursor Y on the canvas (clamped to slide bounds). */
    cursorY: number;
    /** ISO timestamp of last update (for stale-presence cleanup). */
    lastUpdated: string;
    /** Optional currently selected element ID. */
    selectedElementId?: string;
    /** Role in the session (broadcaster, viewer, or collaborator). */
    role?: CollaborationRole;
}
/** Value exposed by `CollaborationContext`. */
interface CollaborationContextValue {
    /** Current WebSocket connection status. */
    status: ConnectionStatus;
    /** Presence data for all remote users (excludes the local user). */
    remoteUsers: UserPresence[];
    /** Broadcast the local user's presence state. */
    broadcastPresence: (update: Partial<Omit<UserPresence, 'clientId'>>) => void;
    /** Total number of connected users (including local). */
    connectedCount: number;
    /** The collaboration config that was provided. */
    config: CollaborationConfig;
    /** The Yjs document (for document sync). */
    doc: Doc | null;
    /** Manually retry the WebSocket connection after a timeout or error. */
    retry: () => void;
}

/**
 * Base handle interface for file viewer components.
 * Defined locally to avoid dependency on external file-viewer plugin packages.
 * Provides a standard `getContent` method used by the host application to
 * retrieve the current file content (e.g. for saving).
 */
interface FileViewerHandle {
    /** Get the current content of the file (for saving) */
    getContent: () => Promise<string | Uint8Array>;
}
/** A shape preset entry used in the toolbar shape insertion palette. */
interface ShapePreset {
    type: SupportedShapeType;
    label: string;
    icon: React__default.ReactNode;
}
/** Tracks the position and target element of an open context menu. */
interface ElementContextMenuState {
    x: number;
    y: number;
    elementId: string;
}
/** Identifies an action triggered from the element right-click context menu. */
type ElementContextMenuAction = 'copy' | 'cut' | 'paste' | 'duplicate' | 'delete' | 'bring-forward' | 'send-backward' | 'bring-front' | 'send-back' | 'bringForward' | 'sendBackward' | 'bringToFront' | 'sendToBack' | 'comment' | 'addComment' | 'group' | 'ungroup' | 'editPoints' | 'editHyperlink';
/**
 * State of an active marquee (rubber-band) selection rectangle.
 * Created when the user clicks and drags on the canvas background,
 * and used to compute which elements fall within the selection area.
 */
interface MarqueeSelectionState {
    startX: number;
    startY: number;
    currentX: number;
    currentY: number;
    additive: boolean;
    baseSelectionIds?: string[];
}
/** Tracks which table cell is selected and/or actively being edited. */
interface TableCellEditorState {
    rowIndex: number;
    columnIndex: number;
    /** When true the cell has an active text input. */
    isEditing?: boolean;
    /** Optional multi-cell selection (Shift+Click range). Each entry is {row, col}. */
    selectedCells?: Array<{
        row: number;
        col: number;
    }>;
}
/** A single parsed table cell with its position, text content, and computed CSS style. */
interface ParsedTableCell {
    rowIndex: number;
    columnIndex: number;
    text: string;
    style: React__default.CSSProperties;
    rawCell: XmlObject;
}
/** Complete parsed table data including row/column structure and all cells. */
interface ParsedTableData {
    rowCount: number;
    columnCount: number;
    rows: XmlObject[];
    columnPercentages: number[];
    cells: ParsedTableCell[];
}
/** Runtime state for a single element's animation during presentation mode. */
interface PresentationAnimationRuntime {
    elementId: string;
    state: 'hidden' | 'entering' | 'visible';
    animation: PptxElementAnimation;
}
/**
 * Groups slides into named sections for the slides pane sidebar.
 * Corresponds to OOXML `p15:section` elements in `presentation.xml`.
 */
interface SlideSectionGroup {
    id: string;
    label: string;
    slideIndexes: number[];
    /** Section highlight color from p15:sectionPr. */
    color?: string;
    /** Whether the section should start collapsed (from p15:sectionPr). */
    defaultCollapsed?: boolean;
}
/** Alignment direction for distributing/aligning multiple selected elements on the slide. */
type SlideAlignment = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom';
/** Identifies one of the ribbon-style toolbar tabs (home, insert, text, etc.). */
type ToolbarSection = 'file' | 'home' | 'insert' | 'text' | 'arrange' | 'draw' | 'design' | 'transitions' | 'animations' | 'slideShow' | 'review' | 'view' | 'help';
/** The active drawing/inking tool selected in the Draw toolbar tab. */
type DrawingTool = 'select' | 'pen' | 'highlighter' | 'eraser' | 'freeform';
/** A single entry in the keyboard shortcuts help panel. */
interface ShortcutReferenceItem {
    action: string;
    shortcut: string;
}
/** An accessibility audit finding (missing alt text, reading order issues, etc.). */
interface AccessibilityIssue {
    slideIndex: number;
    elementId?: string;
    severity: 'error' | 'warning' | 'info';
    message: string;
}
/** Dropdown option for selecting a connector geometry type. */
interface ConnectorGeometryOption {
    value: ConnectorGeometryType;
    label: string;
}
/** Dropdown option for selecting a connector arrowhead style. */
interface ConnectorArrowOption {
    value: ConnectorArrowType;
    label: string;
}
/** Dropdown option for selecting a stroke dash pattern. */
interface StrokeDashOption {
    value: StrokeDashType;
    label: string;
}
/** Dropdown option for selecting a slide transition type. */
interface SlideTransitionOption {
    value: NonNullable<PptxSlideTransition['type']>;
    label: string;
}
/** Dropdown option for selecting an animation effect preset. */
interface AnimationPresetOption {
    value: Exclude<PptxAnimationPreset, 'none'>;
    label: string;
}
interface PowerPointViewerProps {
    /** PowerPoint content as Uint8Array */
    content: Uint8Array;
    /** Original file path — used for autosave recovery */
    filePath?: string;
    /** Callback when content has unsaved changes */
    onDirtyChange?: (isDirty: boolean) => void;
    onContentChange?: (content: Uint8Array) => void;
    /** Callback when active slide changes */
    onActiveSlideChange?: (slideIndex: number) => void;
    /** Whether editing actions are enabled */
    canEdit?: boolean;
    /** Optional class name */
    className?: string;
    /**
     * Display name used as the author for comments and annotations.
     * Falls back to `collaboration.userName` when collaborating, or `'You'`.
     */
    authorName?: string;
    /**
     * Theme configuration for customising the viewer's appearance.
     *
     * Accepts partial color overrides, a custom border-radius, and
     * arbitrary CSS custom properties. Unset values fall back to the
     * built-in dark theme.
     *
     * @example
     * ```tsx
     * <PowerPointViewer
     *   content={bytes}
     *   theme={{
     *     colors: { primary: "#6366f1", background: "#0f172a" },
     *     radius: "0.75rem",
     *   }}
     * />
     * ```
     *
     * @see {@link ViewerTheme} for the full type definition.
     */
    theme?: ViewerTheme;
    /**
     * Optional real-time collaboration configuration.
     *
     * When provided, the viewer enables collaborative editing with live
     * cursors, user presence indicators, and CRDT-based state sync via Yjs.
     * Requires `yjs` and `y-websocket` peer dependencies.
     *
     * @example
     * ```tsx
     * <PowerPointViewer
     *   content={bytes}
     *   collaboration={{
     *     roomId: "my-room-123",
     *     serverUrl: "wss://collab.example.com",
     *     userName: "Alice",
     *     userColor: "#6366f1",
     *   }}
     * />
     * ```
     */
    collaboration?: CollaborationConfig;
    /**
     * Callback invoked when the user starts a collaboration session from the
     * Share dialog. The host app should use this to set the `collaboration`
     * prop with the returned config.
     */
    onStartCollaboration?: (config: CollaborationConfig) => void;
    /**
     * Callback invoked when the user stops a collaboration session from the
     * Share dialog. The host app should clear the `collaboration` prop.
     */
    onStopCollaboration?: () => void;
    /**
     * Default values for the Share dialog fields. The host app should provide
     * these to control the session name, user display name, and server URL.
     * If omitted, the Share dialog fields will be empty and require user input.
     *
     * @example
     * ```tsx
     * <PowerPointViewer
     *   shareDefaults={{
     *     roomId: "session-abc123",
     *     userName: "Alice",
     *     serverUrl: "ws://localhost:1234",
     *   }}
     * />
     * ```
     */
    shareDefaults?: {
        roomId?: string;
        userName?: string;
        serverUrl?: string;
    };
}
interface PowerPointViewerHandle extends FileViewerHandle {
    getContent: () => Promise<Uint8Array>;
}

declare function getAnimationInitialStyle(preset: PptxAnimationPreset | undefined, nativeAnimation?: PptxNativeAnimation): React__default.CSSProperties;

/**
 * Root React component for the PowerPoint viewer/editor.
 *
 * Accepts binary `.pptx` content and renders a full-featured editor with
 * slide canvas, toolbar, inspector panels, presentation mode, and more.
 *
 * Uses `forwardRef` to expose a `PowerPointViewerHandle` for imperative
 * access (e.g. serialising the current content for saving).
 */
declare const PowerPointViewer: React.ForwardRefExoticComponent<PowerPointViewerProps & React.RefAttributes<PowerPointViewerHandle>>;

export { type AccessibilityIssue as A, type SlideAlignment as B, type CollaborationConfig as C, type DragState as D, type EditorHistorySnapshot as E, type FileViewerHandle as F, type SlideSectionGroup as G, type SlideTransitionOption as H, type StrokeDashOption as I, type SupportedShapeType as J, type ToolbarSection as K, type ViewerMode as L, type MarqueeSelectionState as M, PowerPointViewer as P, type ResizeHandle as R, type ShapeAdjustmentDragState as S, type TableCellEditorState as T, type UserPresence as U, type ViewerThemeColors as V, type ViewerTheme as a, type PowerPointViewerHandle as b, type PowerPointViewerProps as c, type ConnectionStatus as d, type CollaborationContextValue as e, type AnimationPresetOption as f, getAnimationInitialStyle as g, type CanvasSize as h, type CollaborationRole as i, type ConnectorArrowOption as j, type ConnectorGeometryOption as k, type ConnectorGeometryType as l, type ConnectorPathGeometry as m, type DrawingTool as n, type ElementBounds as o, type ElementClipboardPayload as p, type ElementContextMenuAction as q, type ElementContextMenuState as r, type ParsedTableCell as s, type ParsedTableData as t, type PresentationAnimationRuntime as u, type ResizeState as v, type ShapeAdjustmentHandleDescriptor as w, type ShapePreset as x, type ShapeQuickStyle as y, type ShortcutReferenceItem as z };
