export interface OfficeWeaverRuntimeInfo {
  engine: string;
  version: string;
  kind: string;
  versionFamily: string;
}

export interface OfficeWeaverBuildOptions {
  engine?: 'onlyoffice' | string;
  version?: string;
  kind?: 'spreadsheet' | 'word' | 'text-document' | string;
}

export interface OfficeWeaverApi {
  version: string;
  resolveVersionFamily(engine?: string, version?: string): string;
  normalizeSpreadsheetMacroCode(codeText?: string): string;
  normalizeTextDocumentMacroCode(codeText?: string): string;
  buildSpreadsheetMacroCommand(codeText: string, options?: OfficeWeaverBuildOptions): Function;
  buildTextDocumentMacroCommand(codeText: string, options?: OfficeWeaverBuildOptions): Function;
  buildWordMacroCommand(codeText: string, options?: OfficeWeaverBuildOptions): Function;
  createRuntime(options?: OfficeWeaverBuildOptions): OfficeWeaverRuntimeInfo;
}

declare const OfficeWeaver: OfficeWeaverApi;

export = OfficeWeaver;
