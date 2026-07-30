export interface FindResult {
  matchIndex: number;
}

/** Search surface shared by DOCX, PPTX, and XLSX viewers. */
export interface FindableViewer {
  findText(
    query: string,
    opts?: { caseSensitive?: boolean },
  ): Promise<FindResult[]>;
  findNext(): Promise<FindResult | null>;
  findPrev(): Promise<FindResult | null>;
  clearFind(): void;
}

export interface FindPopupElements {
  root: HTMLElement;
  input: HTMLInputElement;
  status: HTMLElement;
  previous: HTMLButtonElement;
  next: HTMLButtonElement;
  close: HTMLButtonElement;
}

type KeyTarget = Pick<EventTarget, 'addEventListener' | 'removeEventListener'>;

/**
 * VS Code webviews do not expose the editor's native find widget to custom
 * editors. This controller supplies browser-style find semantics over the
 * viewers' format-aware search APIs.
 */
export class FindPopupController {
  private viewer: FindableViewer | null = null;
  private request = 0;
  private total = 0;
  private chain: Promise<void> = Promise.resolve();
  private readonly onInput = () => {
    void this.search();
  };
  private readonly onPrevious = () => {
    void this.navigate(true);
  };
  private readonly onNext = () => {
    void this.navigate(false);
  };
  private readonly onClose = () => this.close();
  private readonly onKeyDown = (event: Event) => {
    this.handleKeyDown(event as KeyboardEvent);
  };

  constructor(
    private readonly elements: FindPopupElements,
    private readonly keyTarget: KeyTarget,
    private readonly onError: (error: unknown) => void = () => {},
  ) {
    elements.input.addEventListener('input', this.onInput);
    elements.previous.addEventListener('click', this.onPrevious);
    elements.next.addEventListener('click', this.onNext);
    elements.close.addEventListener('click', this.onClose);
    keyTarget.addEventListener('keydown', this.onKeyDown, true);
    this.updateStatus();
  }

  setViewer(viewer: FindableViewer | null): void {
    if (viewer === this.viewer) return;
    this.viewer?.clearFind();
    this.viewer = viewer;
    this.request++;
    this.total = 0;
    this.updateStatus();
    if (!this.elements.root.hidden && this.elements.input.value.length > 0) {
      void this.search();
    }
  }

  open(): void {
    const wasHidden = this.elements.root.hidden;
    this.elements.root.hidden = false;
    this.elements.input.focus();
    this.elements.input.select();
    if (wasHidden && this.elements.input.value.length > 0) void this.search();
  }

  close(): void {
    if (this.elements.root.hidden) return;
    this.elements.root.hidden = true;
    this.request++;
    this.total = 0;
    this.viewer?.clearFind();
    this.updateStatus();
  }

  search(): Promise<void> {
    const request = ++this.request;
    const viewer = this.viewer;
    const query = this.elements.input.value;
    return this.enqueue(async () => {
      if (request !== this.request) return;
      if (!viewer || query.length === 0) {
        viewer?.clearFind();
        this.total = 0;
        this.updateStatus();
        return;
      }

      // Explicitly request case-insensitive matching for all three formats.
      const matches = await viewer.findText(query, { caseSensitive: false });
      if (request !== this.request || viewer !== this.viewer) {
        if (this.elements.root.hidden && viewer === this.viewer) viewer.clearFind();
        return;
      }
      this.total = matches.length;
      if (this.total === 0) {
        this.updateStatus();
        return;
      }
      const active = await viewer.findNext();
      if (request !== this.request || viewer !== this.viewer) return;
      this.updateStatus(active?.matchIndex ?? 0);
    });
  }

  navigate(previous: boolean): Promise<void> {
    const request = this.request;
    const viewer = this.viewer;
    return this.enqueue(async () => {
      if (
        request !== this.request ||
        !viewer ||
        viewer !== this.viewer ||
        this.total === 0
      ) return;
      const active = previous ? await viewer.findPrev() : await viewer.findNext();
      if (request !== this.request || viewer !== this.viewer) return;
      this.updateStatus(active?.matchIndex ?? -1);
    });
  }

  dispose(): void {
    this.close();
    this.elements.input.removeEventListener('input', this.onInput);
    this.elements.previous.removeEventListener('click', this.onPrevious);
    this.elements.next.removeEventListener('click', this.onNext);
    this.elements.close.removeEventListener('click', this.onClose);
    this.keyTarget.removeEventListener('keydown', this.onKeyDown, true);
    this.viewer = null;
  }

  private handleKeyDown(event: KeyboardEvent): void {
    const key = event.key.toLowerCase();
    if ((event.ctrlKey || event.metaKey) && !event.altKey && key === 'f') {
      event.preventDefault();
      event.stopPropagation();
      this.open();
      return;
    }
    if (this.elements.root.hidden) return;
    if (key === 'escape') {
      event.preventDefault();
      event.stopPropagation();
      this.close();
      return;
    }
    if (key === 'enter') {
      const target = event.target;
      if (
        target === this.elements.previous ||
        target === this.elements.next ||
        target === this.elements.close
      ) return;
      event.preventDefault();
      event.stopPropagation();
      void this.navigate(event.shiftKey);
    }
  }

  private enqueue(task: () => Promise<void>): Promise<void> {
    const result = this.chain.then(task);
    this.chain = result.catch(this.onError);
    return result;
  }

  private updateStatus(activeIndex = -1): void {
    this.elements.previous.disabled = this.total === 0;
    this.elements.next.disabled = this.total === 0;
    if (this.elements.input.value.length === 0) {
      this.elements.status.textContent = '';
    } else if (this.total === 0) {
      this.elements.status.textContent = 'No results';
    } else {
      const current = Math.min(this.total, Math.max(1, activeIndex + 1));
      this.elements.status.textContent = `${current} of ${this.total}`;
    }
  }
}
