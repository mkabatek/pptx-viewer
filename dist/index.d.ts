import { V as ViewerThemeColors, a as ViewerTheme } from './PowerPointViewer-C5jGuKGB.js';
export { P as PowerPointViewer, b as PowerPointViewerHandle, c as PowerPointViewerProps, g as getAnimationInitialStyle } from './PowerPointViewer-C5jGuKGB.js';
import { Options } from 'html2canvas-pro';
import * as react_jsx_runtime from 'react/jsx-runtime';
import 'pptx-viewer-core';
import 'react';
import 'yjs';

/**
 * Default dark-theme color values.
 *
 * These correspond to the built-in dark UI of the PowerPoint viewer and
 * use Tailwind's gray palette as the neutral scale with indigo as the
 * primary accent.
 */
declare const defaultThemeColors: ViewerThemeColors;
/** Default border-radius. */
declare const defaultRadius = "0.5rem";

/**
 * Convert a `ViewerTheme` into a flat `Record<string, string>` of CSS
 * custom properties (including the `--` prefix) ready to be spread onto
 * a `style` attribute.
 *
 * Only properties that differ from the built-in defaults are emitted when
 * `omitDefaults` is true (the default).
 */
declare function themeToCssVars(theme: ViewerTheme | undefined, omitDefaults?: boolean): Record<string, string>;
/**
 * Build the complete set of CSS custom properties with all defaults.
 * Useful for generating a full fallback stylesheet.
 */
declare function defaultCssVars(): Record<string, string>;

interface ViewerThemeProviderProps {
    theme?: ViewerTheme;
    children: React.ReactNode;
}
/**
 * Provides a `ViewerTheme` to all descendant viewer components.
 *
 * Typically you do **not** need to use this directly — passing a `theme`
 * prop to `<PowerPointViewer>` is sufficient. This provider is exposed
 * for advanced use-cases where you want to wrap multiple viewers or
 * share a theme across a wider subtree.
 */
declare function ViewerThemeProvider({ theme, children }: ViewerThemeProviderProps): react_jsx_runtime.JSX.Element;
/**
 * Returns the active `ViewerTheme` (if any) from the nearest
 * `ViewerThemeProvider`.
 */
declare function useViewerTheme(): ViewerTheme | undefined;

/**
 * A drop-in replacement for `html2canvas(element, options)` that first
 * resolves any oklch / oklab / lch / lab / color() values in the cloned
 * DOM to rgb()/hex, preventing parse errors in html2canvas ≤ 1.x.
 *
 * Three-pronged approach:
 * 1. Patch `<style>` elements to replace oklch in CSS custom properties.
 * 2. Resolve `:root` / `<body>` inline custom properties.
 * 3. Walk every element and convert computed colour values to sRGB.
 *
 * Usage:
 * ```ts
 * import { renderToCanvas } from "../lib/canvas-export";
 * const canvas = await renderToCanvas(element, { scale: 2 });
 * ```
 */
declare function renderToCanvas(element: HTMLElement, options?: Partial<Options>): Promise<HTMLCanvasElement>;

export { ViewerTheme, ViewerThemeColors, ViewerThemeProvider, defaultCssVars, defaultRadius, defaultThemeColors, renderToCanvas, themeToCssVars, useViewerTheme };
