export type ExtractionType = "hierarchy" | "requirements" | "parameters" | "custom";

export interface ParagraphBlock {
  type: "paragraph";
  text: string;
}

export interface ListBlock {
  type: "list";
  listType: "ordered" | "unordered";
  listItems: string[];
}

export interface TableBlock {
  type: "table";
  tableHeaders: string[];
  tableRows: string[][];
}

export type DocumentBlock = ParagraphBlock | ListBlock | TableBlock;

export interface HierarchicalSection {
  heading: string;
  level: number;
  blocks: DocumentBlock[];
}

export interface HierarchicalResult {
  documentTitle: string;
  sections: HierarchicalSection[];
}

export interface RequirementItem {
  id: string;
  title: string;
  description: string;
  category: string;
  priority: "Hög" | "Medium" | "Låg" | string;
  verificationMethod?: "Test" | "Inspektion" | "Analys" | "Demonstration" | "Ej angivet" | string;
  status?: string;
}

export interface RequirementsResult {
  documentTitle: string;
  version?: string;
  requirements: RequirementItem[];
}

export interface ParameterItem {
  name: string;
  value: string;
  unit?: string;
  tolerance?: string;
  context?: string;
}

export interface ParametersResult {
  documentTitle: string;
  parameters: ParameterItem[];
}

export interface ParseResponse {
  success: boolean;
  fileName: string;
  extractionType: ExtractionType;
  textLength: number;
  rawMarkdown: string;
  result: HierarchicalResult | RequirementsResult | ParametersResult | any;
}
