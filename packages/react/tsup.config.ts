import { defineConfig } from 'tsup';

export default defineConfig((options) => ({
	entry: ['src/index.ts', 'src/viewer/index.ts'],
	format: ['esm', 'cjs'],
	dts: true,
	splitting: false,
	sourcemap: false,
	clean: !options.watch,
	external: [
		'react',
		'react-dom',
		'framer-motion',
		'lucide-react',
		'react-icons',
		'html2canvas-pro',
		'jspdf',
		'jszip',
		'fast-xml-parser',
		'clsx',
		'tailwind-merge',
		'i18next',
		'react-i18next',
	],
	// Bundle workspace-only packages so the published dist is self-contained.
	// These are not on npm — consumers cannot install them separately.
	noExternal: ['pptx-viewer-core', 'emf-converter', 'mtx-decompressor'],
	treeshake: true,
	platform: 'browser',
}));
