// Minimal types for markdown-it@14 (the package ships no declarations). Only the
// surface the Markdown importer uses is declared.
declare module "markdown-it" {
  interface Token {
    type: string;
    tag: string;
    content: string;
    meta: any;
    attrs: Array<[string, string]> | null;
  }

  interface StateInline {
    src: string;
    pos: number;
    posMax: number;
    env: any;
    push(type: string, tag: string, nesting: number): Token;
  }

  interface Ruler<T> {
    before(beforeName: string, ruleName: string, fn: T): void;
    after(afterName: string, ruleName: string, fn: T): void;
  }

  type RenderRule = (tokens: Token[], index: number, options: unknown, env: any, self: unknown) => string;

  interface Options {
    html?: boolean;
    xhtmlOut?: boolean;
    breaks?: boolean;
    linkify?: boolean;
    typographer?: boolean;
  }

  class MarkdownIt {
    constructor(presetName?: "default" | "commonmark" | "zero", options?: Options);
    constructor(options?: Options);
    inline: { ruler: Ruler<(state: StateInline, silent: boolean) => boolean> };
    renderer: { rules: Record<string, RenderRule | undefined> };
    validateLink: (url: string) => boolean;
    render(source: string, env?: any): string;
  }

  export default MarkdownIt;
}
