export interface AnnouncementSection {
  readonly title: string;
  readonly paragraphs: readonly string[];
  readonly bullets?: readonly string[];
  readonly kind?: 'summary';
  readonly examples?: readonly AnnouncementExample[];
}

export interface AnnouncementExample {
  readonly title: string;
  readonly code: string;
}

export interface Announcement {
  readonly slug: string;
  readonly date: string;
  readonly label: 'Upcoming release' | 'Engineering note' | 'Release note';
  readonly version?: string;
  readonly title: string;
  readonly summary: string;
  readonly audience: string;
  readonly sections: readonly AnnouncementSection[];
}

export const announcements: readonly Announcement[] = [
  {
    slug: 'v075-resource-governance',
    date: '2026-08-02',
    label: 'Upcoming release',
    version: 'v0.75',
    title: 'Resource limits, typed failures and metrics for large files',
    summary: 'v0.75 applies default inflated-package limits to every DOCX, XLSX and PPTX load, reports measured limit failures with typed errors, and exposes content-free usage metrics.',
    audience: 'Applications that load user-supplied DOCX, XLSX or PPTX files, especially those that customize maxZipEntryBytes or accept unusually large files.',
    sections: [
      {
        title: 'In short',
        kind: 'summary',
        paragraphs: [
          'Most applications do not need to change how they construct a Viewer: omitting resourceLimits selects the standard policy. Applications should, however, treat a typed limit error as an intentional refusal to preview the file rather than as an unknown renderer failure.',
        ],
        bullets: [
          'No resourceLimits today: no configuration change is required. v0.75 supplies the defaults.',
          'User-supplied files: catch OoxmlResourceLimitError and OoxmlDecodedImageLimitError and show a clear “too large to preview” result.',
          'Using maxZipEntryBytes: it remains a deprecated compatibility alias, but migrate to resourceLimits.maxArchiveEntryBytes.',
          'Need different limits: collect OoxmlResourceMetrics from representative production files before choosing values.',
        ],
      },
      {
        title: 'What is now bounded',
        paragraphs: [
          'DOCX, XLSX and PPTX now use the same admission policy while opening and lazily reading the ZIP package. The limits apply to inflated package parts—not to the compressed upload size and not to JavaScript heap usage. A part is charged by the largest amount read from it, so reading the same part again does not consume the distinct-total budget twice.',
          'Raster image decoding has separate, non-configurable browser guards. These are hard implementation ceilings because decoded memory and browser/GPU overhead do not map consistently to an application-supplied byte value across devices.',
        ],
        bullets: [
          '128 MiB for any one inflated XML, image, media or other package part by default.',
          '256 MiB across distinct inflated parts read during one package session by default.',
          '32 megapixels per raster image, 128 MiB aggregate decoded raster ownership, and two concurrent image decodes.',
          'Internal hard ceilings still apply when either configurable package limit is set to null.',
        ],
      },
      {
        title: 'Handle an intentional rejection',
        paragraphs: [
          'Catch limit errors wherever your application awaits load or later lazy document work. OoxmlResourceLimitError includes the measured limit and observed value in details.violation; OoxmlDecodedImageLimitError identifies a raster-image guard. The same classes are re-exported by the DOCX, XLSX and PPTX entry points.',
        ],
        examples: [
          {
            title: 'Show a specific preview error',
            code: `import {
  DocxViewer,
  OoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
} from '@silurus/ooxml/docx';

const viewer = new DocxViewer(canvas);

try {
  await viewer.load(file);
} catch (error) {
  if (error instanceof OoxmlResourceLimitError) {
    const { limit, observed } = error.details.violation;
    showPreviewError(
      \`This file exceeds the preview limit (\${observed} of \${limit} bytes).\`,
    );
    return;
  }

  if (error instanceof OoxmlDecodedImageLimitError) {
    showPreviewError('This file contains an image that is too large to preview.');
    return;
  }

  throw error;
}`,
          },
        ],
      },
      {
        title: 'Choose limits from observed files',
        paragraphs: [
          'Start with the defaults rather than guessing. onResourceMetrics receives a content-free report when the initial load settles, including failed loads. After a successful load and any lazy page, sheet, slide, image or media access, getResourceMetrics() returns a fresh snapshot. The library does not transmit or persist either report.',
          'Metrics exclude filenames, URLs, package paths, document text, passwords and raw error messages. Sizes, counts and timings are still document-derived metadata, so collect them only under your application’s consent and retention policy. debug: true is a separate development aid that prints the same class of data to the console.',
        ],
        examples: [
          {
            title: 'Collect metrics without console output',
            code: `const viewer = new DocxViewer(canvas, {
  onResourceMetrics(metrics) {
    usageMetrics.record(metrics);
  },
});

await viewer.load(file);

// Includes package work observed after the initial load.
usageMetrics.record(await viewer.getResourceMetrics());`,
          },
          {
            title: 'Apply values chosen from your own data',
            code: `const MiB = 1024 * 1024;

const viewer = new DocxViewer(canvas, {
  resourceLimits: {
    maxArchiveEntryBytes: 64 * MiB,
    maxTotalInflatedBytes: 192 * MiB,
  },
});`,
          },
        ],
      },
      {
        title: 'Migrate maxZipEntryBytes',
        paragraphs: [
          'A positive maxZipEntryBytes value keeps its existing per-entry meaning in v0.75, so this migration is not required immediately. New code should use resourceLimits. Do not supply conflicting values through both options; that is rejected before parsing begins.',
        ],
        examples: [
          {
            title: 'Before',
            code: `const viewer = new DocxViewer(canvas, {
  maxZipEntryBytes: 64 * 1024 * 1024,
});`,
          },
          {
            title: 'After',
            code: `const viewer = new DocxViewer(canvas, {
  resourceLimits: {
    maxArchiveEntryBytes: 64 * 1024 * 1024,
  },
});`,
          },
        ],
      },
      {
        title: 'What these limits cannot guarantee',
        paragraphs: [
          'Package counters do not measure peak process memory. XML trees, document models, canvas backing stores, decoded images, renderer state and browser-managed memory can require several times the measured inflated bytes. The defaults reject known measurable hazards earlier, but cannot guarantee that every browser and device will avoid an out-of-memory termination.',
          'A residual WebAssembly trap is reported conservatively as parser-crashed, not parser-oom. At the current WASM boundary, Rust panic, allocation failure, stack overflow and explicit unreachable can converge on the same WebAssembly.RuntimeError, so the original cause cannot be recovered reliably after the trap. Worker mode can keep parser and renderer work away from the Window and improve failure containment, but a Worker is not a separate operating-system process or a strict memory sandbox.',
        ],
      },
    ],
  },
];

export const latestAnnouncements = announcements.slice(0, 3);

export function formatAnnouncementDate(value: string): string {
  return new Intl.DateTimeFormat('en', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    timeZone: 'UTC',
  }).format(new Date(`${value}T00:00:00Z`));
}
