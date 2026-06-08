/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, DragEvent, ChangeEvent } from "react";
import { motion, AnimatePresence } from "motion/react";
import {
  UploadCloud,
  FileJson,
  CheckCircle2,
  AlertCircle,
  Loader2,
  Download,
  Copy,
  ArrowLeft,
  Search,
  Settings,
  ListFilter,
  Layers,
  FileText,
  ShieldCheck,
  Tag,
  Sliders,
  Database,
  Eye,
  Menu,
  FileSpreadsheet
} from "lucide-react";
import {
  ExtractionType,
  ParseResponse,
  HierarchicalResult,
  RequirementsResult,
  ParametersResult,
  RequirementItem,
  ParameterItem
} from "./types";

export default function App() {
  // Application State
  const [file, setFile] = useState<File | null>(null);
  const [base64Data, setBase64Data] = useState<string | null>(null);
  const [extractionType, setExtractionType] = useState<ExtractionType>("hierarchy");
  const [customPrompt, setCustomPrompt] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [loadingStep, setLoadingStep] = useState<number>(0);
  const [response, setResponse] = useState<ParseResponse | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  
  // Tab control in Results
  const [activeTab, setActiveTab] = useState<"structured" | "json" | "raw_markdown">("structured");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [priorityFilter, setPriorityFilter] = useState<string>("all");
  const [copied, setCopied] = useState<boolean>(false);

  // Drag-and-drop state
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Dynamic loading phase statements
  const loadingSteps = [
    "Läser in Word-dokumentets råa data...",
    "Extraherar innehåll och omformaterar till strukturerad Markdown...",
    "Ansluter säkert till Gemini 3.5 AI-motorn...",
    "Tolkar formatering, dolda tabeller och rubriknivåer...",
    "Bygger en komplett, oförändrad JSON-fil baserat på det valda schemat...",
    "Slutför sammanställningen och strukturerar de sista fälten..."
  ];

  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const processFile = async (selectedFile: File) => {
    if (!selectedFile.name.endsWith(".docx")) {
      setErrorMsg("Bara Word-filer (.docx) stöds för tillfället.");
      return;
    }
    setErrorMsg(null);
    setFile(selectedFile);

    // Convert file to Base64
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      // Extract the raw base64 string from data URL
      const base64Str = result.split(",")[1];
      setBase64Data(base64Str);
    };
    reader.onerror = () => {
      setErrorMsg("Kunde inte läsa in filen. Försök igen.");
    };
    reader.readAsDataURL(selectedFile);
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processFile(e.target.files[0]);
    }
  };

  const triggerFileSelect = () => {
    fileInputRef.current?.click();
  };

  const handleReset = () => {
    setFile(null);
    setBase64Data(null);
    setResponse(null);
    setErrorMsg(null);
    setIsLoading(false);
    setLoadingStep(0);
    setActiveTab("structured");
    setSearchQuery("");
    setPriorityFilter("all");
  };

  const handleConvert = async () => {
    if (!base64Data || !file) {
      setErrorMsg("Vänligen ladda upp en Word-fil först.");
      return;
    }

    setIsLoading(true);
    setLoadingStep(0);
    setErrorMsg(null);

    // Cycle through loading steps visually while processing
    const stepInterval = setInterval(() => {
      setLoadingStep((prev) => (prev < loadingSteps.length - 1 ? prev + 1 : prev));
    }, 2000);

    try {
      const apiResponse = await fetch("/api/parse-docx", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          fileData: base64Data,
          fileName: file.name,
          extractionType,
          customPrompt: extractionType === "custom" ? customPrompt : undefined,
        }),
      });

      clearInterval(stepInterval);

      if (!apiResponse.ok) {
        const errorData = await apiResponse.json();
        throw new Error(errorData.error || "Misslyckades att tolka dokumentet.");
      }

      const data: ParseResponse = await apiResponse.json();
      setResponse(data);
    } catch (err: any) {
      clearInterval(stepInterval);
      setErrorMsg(err.message || "Ett anslutningsfel inträffade vid kontakt med servern.");
      setIsLoading(false);
    } finally {
      setIsLoading(false);
    }
  };

  const handleCopyJson = () => {
    if (!response) return;
    const jsonStr = JSON.stringify(response.result, null, 2);
    navigator.clipboard.writeText(jsonStr);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleDownloadJson = () => {
    if (!response) return;
    const jsonStr = JSON.stringify(response.result, null, 2);
    const blob = new Blob([jsonStr], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    // Derive name
    const originalName = file ? file.name.replace(/\.[^/.]+$/, "") : "document";
    a.download = `${originalName}_extracted_${extractionType}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Helper parser for nested hierarchical trees
  const renderHierarchySection = (sections: any[]) => {
    return sections.map((sec: any, sIdx: number) => (
      <div key={sIdx} className="mb-8 border-l-2 border-slate-200 pl-4 sm:pl-6 ml-1 py-1">
        <div className="flex items-center gap-2 mb-3">
          <span className="text-xs font-mono px-2 py-0.5 rounded-md bg-indigo-50 text-indigo-700 font-semibold border border-indigo-155">
            H{sec.level}
          </span>
          <h3 className={`font-sans font-bold text-slate-800 uppercase tracking-tight ${
            sec.level === 1 ? "text-xl sm:text-2xl" : sec.level === 2 ? "text-lg sm:text-xl" : "text-base"
          }`}>
            {sec.heading}
          </h3>
        </div>
        
        <div className="space-y-4">
          {sec.blocks?.map((block: any, bIdx: number) => {
            if (block.type === "paragraph") {
              return (
                <p key={bIdx} className="text-sm sm:text-base text-slate-655 leading-relaxed font-sans font-normal">
                  {block.text}
                </p>
              );
            }
            if (block.type === "list") {
              const ListTag = block.listType === "ordered" ? "ol" : "ul";
              return (
                <ListTag key={bIdx} className={`pl-6 text-sm text-slate-655 font-sans ${
                  block.listType === "ordered" ? "list-decimal" : "list-disc"
                } space-y-2`}>
                  {block.listItems?.map((item: string, iIdx: number) => (
                    <li key={iIdx} className="leading-relaxed">
                      {item}
                    </li>
                  ))}
                </ListTag>
              );
            }
            if (block.type === "table") {
              return (
                <div key={bIdx} className="my-4 overflow-x-auto rounded-lg border border-slate-200 shadow-sm max-w-full bg-white">
                  <table className="w-full text-left border-collapse font-sans text-xs sm:text-sm">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200 text-slate-705 font-bold">
                        {block.tableHeaders?.map((header: string, hIdx: number) => (
                          <th key={hIdx} className="px-3 py-2 sm:px-4 sm:py-3 whitespace-nowrap">{header}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-150">
                      {block.tableRows?.map((row: string[], rIdx: number) => (
                        <tr key={rIdx} className="hover:bg-slate-50/50 transition">
                          {row.map((cell: string, cIdx: number) => (
                            <td key={cIdx} className="px-3 py-2 sm:px-4 sm:py-3 text-slate-600 font-normal">{cell}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              );
            }
            return null;
          })}
        </div>
      </div>
    ));
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col font-sans selection:bg-indigo-500/10 selection:text-indigo-900">
      
      {/* Upper Subtle Branding Rail (Professional Polish Detail) */}
      <header className="border-b border-slate-200 bg-white sticky top-0 z-40 px-4 sm:px-6 py-4">
        <div className="max-w-6xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="h-8 w-8 bg-indigo-600 rounded flex items-center justify-center text-white font-bold text-sm shadow-sm shrink-0">
              W
            </div>
            <div>
              <h1 className="text-base sm:text-lg font-semibold tracking-tight text-slate-800 mb-0">
                APAB <span className="text-slate-300 font-normal mx-1">|</span> Word to JSON
              </h1>
              <span className="text-xs text-slate-500 font-light block">Komplett, semantisk dokumentkonverterare</span>
            </div>
          </div>
          <div className="text-xs font-mono text-slate-500 flex items-center gap-2">
            <ShieldCheck className="h-4 w-4 text-indigo-600" />
            <span>Säker kryptering</span>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="flex-1 max-w-6xl w-full mx-auto px-4 sm:px-6 py-8 sm:py-12 flex flex-col justify-start">
        
        <AnimatePresence mode="wait">
          
          {/* STEP 1: INITIAL UPLOAD & CONFIGURATION SCREEN */}
          {!response && !isLoading && (
            <motion.div
              key="setup-screen"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.3 }}
              className="space-y-10"
            >
              
              {/* Intent-focused Hero */}
              <div className="text-center max-w-2xl mx-auto space-y-3">
                <span className="text-xs font-bold uppercase tracking-widest text-indigo-600">Plocka ut godbiten i ett svep</span>
                <h2 className="text-3xl sm:text-4xl font-extrabold tracking-tight text-slate-900 leading-tight">
                  Konvertera Word-dokument till kompletta JSON-filer
                </h2>
                <p className="text-sm sm:text-base text-slate-500 font-light leading-relaxed">
                  Ladda upp ditt dokument och extrahera allt innehåll. Gemini analyserar rubriker, parametrar och krav och returnerar en komplett JSON utan förluster.
                </p>
              </div>

              {/* Upload + Strategy Area Grid */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
                
                {/* File Dropzone & Details (Left side) */}
                <div className="lg:col-span-12 xl:lg:col-span-5 h-full space-y-6">
                  <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-4 shadow-sm">
                    <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                      <span className="h-2 w-2 rounded-full bg-indigo-600 block" />
                      1. Välj Word-fil (.docx)
                    </h3>

                    {/* Drag-and-drop zone */}
                    <div
                      onDragOver={handleDragOver}
                      onDragLeave={handleDragLeave}
                      onDrop={handleDrop}
                      onClick={triggerFileSelect}
                      className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition flex flex-col items-center justify-center gap-3 relative overflow-hidden ${
                        isDragging
                          ? "border-indigo-400 bg-indigo-50/50"
                          : file
                          ? "border-indigo-200 bg-indigo-50/10 hover:border-indigo-300"
                          : "border-slate-200 bg-slate-50/30 hover:border-indigo-200 hover:bg-slate-50/80"
                      }`}
                    >
                      <input
                        type="file"
                        ref={fileInputRef}
                        onChange={handleFileChange}
                        accept=".docx"
                        className="hidden"
                      />

                      {file ? (
                        <div className="space-y-2 z-10 w-full">
                          <div className="h-12 w-12 rounded-xl bg-indigo-50 flex items-center justify-center mx-auto text-indigo-650 border border-indigo-100 shadow-sm">
                            <FileText className="h-6 w-6" />
                          </div>
                          <p className="text-sm font-semibold text-slate-800 line-clamp-1 px-4">{file.name}</p>
                          <p className="text-xs text-slate-500 font-light">
                            {(file.size / 1024).toFixed(1)} KB • Redo att konverteras
                          </p>
                        </div>
                      ) : (
                        <div className="space-y-2 z-10">
                          <UploadCloud className="h-10 w-10 text-indigo-400 mx-auto" strokeWidth={1.5} />
                          <p className="text-sm font-semibold text-slate-700">
                            Dra och släpp din Word-fil här
                          </p>
                          <p className="text-xs text-slate-400 font-light">
                            Eller klicka för att bläddra på datorn (.docx)
                          </p>
                        </div>
                      )}

                      {/* Ripple background effect when dragging */}
                      {isDragging && (
                        <div className="absolute inset-0 bg-indigo-50/50 pointer-events-none transition duration-300" />
                      )}
                    </div>
                  </div>

                  {/* Settings Warning Callout */}
                  <div className="bg-white border border-slate-200 rounded-xl p-4 flex gap-3 text-xs text-slate-600 shadow-sm">
                    <AlertCircle className="h-5 w-5 text-amber-500 shrink-0 mt-0.5" />
                    <div>
                      <p className="font-semibold text-slate-800 mb-0.5">Viktigt om API-nycklar</p>
                      <p className="leading-relaxed font-light text-slate-500">
                        Word to JSON använder Gemini API säkert på servern. Om anropet misslyckas, säkerställ att du lagt till din <code className="text-indigo-600 font-mono bg-indigo-50 px-1 py-0.5 rounded">GEMINI_API_KEY</code> i inställningspanelen (Settings &gt; Secrets).
                      </p>
                    </div>
                  </div>
                </div>

                {/* Extraction Mode Setup (Right side) */}
                <div className="lg:col-span-12 xl:lg:col-span-7 space-y-6">
                  <div className="bg-white border border-slate-200 rounded-2xl p-6 space-y-6 shadow-sm">
                    <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                      <span className="h-2 w-2 rounded-full bg-indigo-600 block" />
                      2. Välj Extraktionsmetod
                    </h3>

                    {/* Card grid for strategies */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      
                      {/* Structure Card 1: Hierarchy */}
                      <div
                        onClick={() => setExtractionType("hierarchy")}
                        className={`p-4 rounded-xl border text-left cursor-pointer transition duration-200 flex flex-col justify-between h-32 ${
                          extractionType === "hierarchy"
                            ? "border-indigo-500 bg-indigo-50/20 shadow-sm shadow-indigo-100"
                            : "border-slate-200 bg-slate-50/30 hover:border-slate-300 hover:bg-slate-50/70"
                        }`}
                      >
                        <div className="flex justify-between items-start">
                          <Layers className={`h-5 w-5 ${extractionType === "hierarchy" ? "text-indigo-600" : "text-slate-400"}`} />
                          <span className={`h-4 w-4 rounded-full border flex items-center justify-center ${
                            extractionType === "hierarchy" ? "border-indigo-600 bg-indigo-600 text-white font-sans" : "border-slate-300 bg-white"
                          }`}>
                            {extractionType === "hierarchy" && <span className="h-1.5 w-1.5 rounded-full bg-white" />}
                          </span>
                        </div>
                        <div>
                          <p className="text-sm font-bold text-slate-800 mb-0.5">Hierarkiskt Träd</p>
                          <p className="text-xs text-slate-500 line-clamp-2 leading-relaxed">
                            Organiserar rubriker, paragrafer, punktlistor och tabeller i läsordning.
                          </p>
                        </div>
                      </div>

                      {/* Structure Card 2: Requirements */}
                      <div
                        onClick={() => setExtractionType("requirements")}
                        className={`p-4 rounded-xl border text-left cursor-pointer transition duration-200 flex flex-col justify-between h-32 ${
                          extractionType === "requirements"
                            ? "border-indigo-500 bg-indigo-50/20 shadow-sm shadow-indigo-100"
                            : "border-slate-200 bg-slate-50/30 hover:border-slate-300 hover:bg-slate-50/70"
                        }`}
                      >
                        <div className="flex justify-between items-start">
                          <ShieldCheck className={`h-5 w-5 ${extractionType === "requirements" ? "text-indigo-600" : "text-slate-400"}`} />
                          <span className={`h-4 w-4 rounded-full border flex items-center justify-center ${
                            extractionType === "requirements" ? "border-indigo-600 bg-indigo-600 text-white font-sans" : "border-slate-300 bg-white"
                          }`}>
                            {extractionType === "requirements" && <span className="h-1.5 w-1.5 rounded-full bg-white" />}
                          </span>
                        </div>
                        <div>
                          <p className="text-sm font-bold text-slate-800 mb-0.5">Kravspecifikation</p>
                          <p className="text-xs text-slate-500 line-clamp-2 leading-relaxed">
                            Extraherar alla krav, standarder eller regler med ID, prioritet och verifiering.
                          </p>
                        </div>
                      </div>

                      {/* Structure Card 3: Parameters */}
                      <div
                        onClick={() => setExtractionType("parameters")}
                        className={`p-4 rounded-xl border text-left cursor-pointer transition duration-200 flex flex-col justify-between h-32 ${
                          extractionType === "parameters"
                            ? "border-indigo-500 bg-indigo-50/20 shadow-sm shadow-indigo-100"
                            : "border-slate-200 bg-slate-50/30 hover:border-slate-300 hover:bg-slate-50/70"
                        }`}
                      >
                        <div className="flex justify-between items-start">
                          <Sliders className={`h-5 w-5 ${extractionType === "parameters" ? "text-indigo-600" : "text-slate-400"}`} />
                          <span className={`h-4 w-4 rounded-full border flex items-center justify-center ${
                            extractionType === "parameters" ? "border-indigo-600 bg-indigo-600 text-white font-sans" : "border-slate-300 bg-white"
                          }`}>
                            {extractionType === "parameters" && <span className="h-1.5 w-1.5 rounded-full bg-white" />}
                          </span>
                        </div>
                        <div>
                          <p className="text-sm font-bold text-slate-800 mb-0.5">Tekniska Parametrar</p>
                          <p className="text-xs text-slate-500 line-clamp-2 leading-relaxed">
                            Sammanställer inställningar, numeriska värven, toleranser och mätetal.
                          </p>
                        </div>
                      </div>

                      {/* Structure Card 4: Custom Structure */}
                      <div
                        onClick={() => setExtractionType("custom")}
                        className={`p-4 rounded-xl border text-left cursor-pointer transition duration-200 flex flex-col justify-between h-32 ${
                          extractionType === "custom"
                            ? "border-indigo-500 bg-indigo-50/20 shadow-sm shadow-indigo-100"
                            : "border-slate-200 bg-slate-50/30 hover:border-slate-300 hover:bg-slate-50/70"
                        }`}
                      >
                        <div className="flex justify-between items-start">
                          <Database className={`h-5 w-5 ${extractionType === "custom" ? "text-indigo-600" : "text-slate-400"}`} />
                          <span className={`h-4 w-4 rounded-full border flex items-center justify-center ${
                            extractionType === "custom" ? "border-indigo-600 bg-indigo-600 text-white font-sans" : "border-slate-300 bg-white"
                          }`}>
                            {extractionType === "custom" && <span className="h-1.5 w-1.5 rounded-full bg-white" />}
                          </span>
                        </div>
                        <div>
                          <p className="text-sm font-bold text-slate-800 mb-0.5">Anpassad Struktur</p>
                          <p className="text-xs text-slate-500 line-clamp-2 leading-relaxed">
                            Skapa din helt unika JSON-layout genom naturliga instruktioner.
                          </p>
                        </div>
                      </div>

                    </div>

                    {/* Inline Text Area for Custom extraction instruction */}
                    <AnimatePresence>
                      {extractionType === "custom" && (
                        <motion.div
                          initial={{ opacity: 0, height: 0 }}
                          animate={{ opacity: 1, height: "auto" }}
                          exit={{ opacity: 0, height: 0 }}
                          transition={{ duration: 0.2 }}
                          className="space-y-2 overflow-hidden"
                        >
                          <label className="text-xs font-bold uppercase tracking-wider text-slate-600 block">
                            Specialinstruktion för anpassat JSON-schema
                          </label>
                          <textarea
                            value={customPrompt}
                            onChange={(e) => setCustomPrompt(e.target.value)}
                            placeholder="T.ex: Skapa en platt lista över alla kontaktpersoner, deras telefonnummer, email och vilken maskindel de ansvarar för."
                            className="w-full h-24 rounded-lg bg-slate-50 border border-slate-200 p-3 text-sm text-slate-800 placeholder:text-slate-400 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition"
                          />
                        </motion.div>
                      )}
                    </AnimatePresence>

                    {/* Conversion Launch Button */}
                    <button
                      onClick={handleConvert}
                      disabled={!file}
                      className={`w-full py-4 rounded-xl font-bold tracking-wide transition shadow-lg flex items-center justify-center gap-2 ${
                        file
                          ? "bg-indigo-600 hover:bg-indigo-700 text-white shadow-lg shadow-indigo-100 cursor-pointer text-base active:scale-[0.98]"
                          : "bg-slate-200 text-slate-400 shadow-none cursor-not-allowed text-base"
                      }`}
                    >
                      <FileJson className="h-5 w-5" />
                      Extrahera komplett JSON
                    </button>

                  </div>
                </div>

              </div>

              {/* Error boundary placement */}
              {errorMsg && (
                <div className="bg-rose-50 border border-rose-200 text-rose-700 p-4 rounded-xl flex gap-3 text-sm leading-relaxed items-center shadow-sm">
                  <AlertCircle className="h-5 w-5 text-rose-500 shrink-0" />
                  <div className="flex-1 font-sans">{errorMsg}</div>
                </div>
              )}

            </motion.div>
          )}

          {/* STEP 2: LOADING SCREEN */}
          {isLoading && (
            <motion.div
              key="loading-screen"
              initial={{ opacity: 0, scale: 0.98 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 1.02 }}
              className="max-w-md mx-auto py-20 text-center flex flex-col items-center justify-center gap-6"
            >
              <div className="relative">
                {/* Glowing ring animation */}
                <div className="absolute inset-0 rounded-full bg-indigo-500/10 blur-xl animate-pulse" />
                <div className="h-20 w-20 rounded-2xl bg-white border border-slate-200 shadow-md flex items-center justify-center relative">
                  <Loader2 className="h-10 w-10 text-indigo-600 animate-spin" />
                </div>
              </div>

              <div className="space-y-2">
                <h3 className="text-xl font-bold text-slate-800 tracking-tight">Formaterar allt innehåll</h3>
                <p className="text-xs text-indigo-600 font-bold tracking-widest uppercase font-mono">Steg {loadingStep + 1} av {loadingSteps.length}</p>
              </div>

              {/* Carousel of loader statements */}
              <div className="h-12 flex items-center justify-center w-full px-4">
                <AnimatePresence mode="wait">
                  <motion.p
                    key={loadingStep}
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -10 }}
                    transition={{ duration: 0.2 }}
                    className="text-sm text-slate-550 font-medium"
                  >
                    {loadingSteps[loadingStep]}
                  </motion.p>
                </AnimatePresence>
              </div>

              {/* Progress bar indicator */}
              <div className="w-full bg-slate-200 h-2 rounded-full overflow-hidden border border-slate-300 shadow-inner">
                <motion.div
                  className="bg-indigo-600 h-full rounded-full"
                  initial={{ width: "0%" }}
                  animate={{ width: `${((loadingStep + 1) / loadingSteps.length) * 100}%` }}
                  transition={{ duration: 0.5 }}
                />
              </div>

              <p className="text-xs text-slate-555 leading-relaxed font-light mt-6 max-w-xs block">
                Tack vare Gemini 3.5 semantisk djup-läsning hanteras tabeller, paragrafer och indrag med precision. Detta tar vanligtvis under 10 sekunder.
              </p>
            </motion.div>
          )}

          {/* STEP 3: HIGH FIDELITY OUTPUT DASHBOARD */}
          {response && !isLoading && (
            <motion.div
              key="result-screen"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.4 }}
              className="space-y-8"
            >
              
              {/* Back button and quick operational summary */}
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-200 pb-6">
                <div className="flex items-center gap-3">
                  <button
                    onClick={handleReset}
                    className="h-10 w-10 border border-slate-250 hover:border-slate-350 bg-white hover:bg-slate-50 rounded-xl flex items-center justify-center text-slate-600 hover:text-slate-900 transition cursor-pointer shadow-sm"
                  >
                    <ArrowLeft className="h-4 w-4" />
                  </button>
                  <div>
                    <h2 className="text-xl font-extrabold tracking-tight text-slate-900 mb-0">{file?.name}</h2>
                    <p className="text-xs text-slate-550 font-light flex items-center gap-2 mt-1">
                      <span className="text-indigo-6 font-bold bg-indigo-50 border border-indigo-100 px-2 py-0.5 rounded">100% Komplett Extraktion</span>
                      <span>•</span>
                      <span className="font-semibold text-slate-600">{response.extractionType.toUpperCase()}</span>
                      <span>•</span>
                      <span>{(response.textLength / 1000).toFixed(1)}k ord-tecken</span>
                    </p>
                  </div>
                </div>

                {/* Main Dashboard Download & Copy Controls */}
                <div className="flex items-center gap-3">
                  <button
                    onClick={handleCopyJson}
                    className="px-4 py-2 text-sm border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 hover:text-slate-900 rounded-xl transition flex items-center gap-2 cursor-pointer font-semibold shadow-sm"
                  >
                    {copied ? (
                      <>
                        <CheckCircle2 className="h-4 w-4 text-emerald-650" />
                        <span className="text-emerald-700">Kopierat!</span>
                      </>
                    ) : (
                      <>
                        <Copy className="h-4 w-4 text-slate-550" />
                        <span>Kopiera JSON</span>
                      </>
                    )}
                  </button>
                  <button
                    onClick={handleDownloadJson}
                    className="px-5 py-2 text-sm bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl transition flex items-center gap-2 shadow-lg shadow-indigo-100 cursor-pointer"
                  >
                    <Download className="h-4 w-4" />
                    <span>Ladda ned JSON-fil</span>
                  </button>
                </div>
              </div>

              {/* Dashboard Layout: Tabs Selector */}
              <div className="flex border-b border-slate-200 gap-6">
                <button
                  onClick={() => setActiveTab("structured")}
                  className={`pb-3 text-sm font-bold tracking-wide border-b-2 transition cursor-pointer ${
                    activeTab === "structured"
                      ? "border-indigo-600 text-indigo-600"
                      : "border-transparent text-slate-500 hover:text-indigo-600"
                  }`}
                >
                  Strukturerat Innehåll
                </button>
                <button
                  onClick={() => setActiveTab("json")}
                  className={`pb-3 text-sm font-bold tracking-wide border-b-2 transition cursor-pointer ${
                    activeTab === "json"
                      ? "border-indigo-600 text-indigo-600"
                      : "border-transparent text-slate-500 hover:text-indigo-600"
                  }`}
                >
                  Rå JSON Kod ({Object.keys(response.result).length} rötter)
                </button>
                <button
                  onClick={() => setActiveTab("raw_markdown")}
                  className={`pb-3 text-sm font-bold tracking-wide border-b-2 transition cursor-pointer ${
                    activeTab === "raw_markdown"
                      ? "border-indigo-600 text-indigo-600"
                      : "border-transparent text-slate-500 hover:text-indigo-600"
                  }`}
                >
                  Inläst Text från Word-filen
                </button>
              </div>

              {/* TAB VIEWS CONTAINER */}
              <div className="min-h-[400px]">
                
                {/* 1. STRUCTURED RENDERER TAB */}
                {activeTab === "structured" && (
                  <motion.div
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    className="space-y-6"
                  >
                    
                    {/* SUB-VIEW A: HIERARCHY */}
                    {extractionType === "hierarchy" && (
                      <div className="bg-white border border-slate-200 rounded-2xl p-6 sm:p-8 space-y-6 shadow-sm">
                        <div className="border-b border-slate-100 pb-4 mb-4">
                          <span className="text-xs text-slate-450 font-mono block mb-1">DOKUMENTTITEL</span>
                          <h2 className="text-2xl sm:text-3xl font-extrabold tracking-tight text-slate-800 font-sans uppercase">
                            {(response.result as HierarchicalResult).documentTitle || "Otitlat Dokument"}
                          </h2>
                        </div>
                        <div className="mt-8">
                          {Array.isArray((response.result as HierarchicalResult).sections) ? (
                            renderHierarchySection((response.result as HierarchicalResult).sections)
                          ) : (
                            <p className="text-slate-500 italic">Hierarki-data saknar sektioner eller är felformaterad.</p>
                          )}
                        </div>
                      </div>
                    )}

                    {/* SUB-VIEW B: REQUIREMENTS LEDGER */}
                    {extractionType === "requirements" && (
                      <div className="space-y-6">
                        
                        {/* Summary Block & Filtering Controls */}
                        <div className="bg-white border border-slate-200 p-6 rounded-2xl flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-sm">
                          <div>
                            <span className="text-xs text-slate-400 font-bold block">DOKUMENTTITEL</span>
                            <h3 className="text-lg font-bold text-slate-800 mb-0 font-sans">
                              {(response.result as RequirementsResult).documentTitle || "Kravspecifikation"}
                            </h3>
                            { (response.result as RequirementsResult).version && (
                              <span className="text-xs font-mono bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded border border-indigo-100/60 mt-1.5 block w-fit">
                                Version {(response.result as RequirementsResult).version}
                              </span>
                            )}
                          </div>

                          <div className="flex flex-wrap items-center gap-3">
                            {/* Filter by Priority */}
                            <div className="flex items-center gap-2">
                              <span className="text-xs text-slate-500 font-medium shrink-0">Prioritet:</span>
                              <select
                                value={priorityFilter}
                                onChange={(e) => setPriorityFilter(e.target.value)}
                                className="bg-slate-50 border border-slate-200 text-xs text-slate-700 outline-none rounded-lg py-1.5 px-3 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 transition"
                              >
                                <option value="all">Alla</option>
                                <option value="hög">Hög / Must / Shall</option>
                                <option value="medium">Medium / Should</option>
                                <option value="låg">Låg / May</option>
                              </select>
                            </div>

                            {/* Search and filter requirements */}
                            <div className="relative">
                              <Search className="h-4 w-4 text-slate-400 absolute left-3 top-2.5" />
                              <input
                                type="text"
                                placeholder="Sök kravtext..."
                                value={searchQuery}
                                onChange={(e) => setSearchQuery(e.target.value)}
                                className="bg-slate-50 border border-slate-200 text-xs text-slate-700 outline-none rounded-lg py-2 pl-9 pr-4 w-52 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 focus:w-64 transition-all"
                              />
                            </div>
                          </div>
                        </div>

                        {/* List requirements as cards */}
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          {(() => {
                            const reqList = (response.result as RequirementsResult).requirements || [];
                            const filtered = reqList.filter((r) => {
                              const matchesSearch = r.id.toLowerCase().includes(searchQuery.toLowerCase()) ||
                                r.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
                                r.description.toLowerCase().includes(searchQuery.toLowerCase()) ||
                                r.category.toLowerCase().includes(searchQuery.toLowerCase());
                              
                              if (priorityFilter === "all") return matchesSearch;
                              if (priorityFilter === "hög") return matchesSearch && ["hög", "high", "skall", "must", "shall"].includes(r.priority.toLowerCase());
                              if (priorityFilter === "medium") return matchesSearch && ["medium", "bör", "should"].includes(r.priority.toLowerCase());
                              if (priorityFilter === "låg") return matchesSearch && ["låg", "låg/lägre", "low", "kan", "may"].includes(r.priority.toLowerCase());
                              
                              return matchesSearch;
                            });

                            if (filtered.length === 0) {
                              return (
                                <div className="col-span-1 md:col-span-2 text-center p-12 border border-slate-200 bg-white rounded-2xl text-slate-400">
                                  Inga krav matchar den valda filtreringen.
                                </div>
                              );
                            }

                            return filtered.map((req: RequirementItem, rIdx: number) => {
                              // Derive priority color tags
                              const prioLower = req.priority.toLowerCase();
                              const isHög = ["hög", "high", "must", "shall", "skall"].includes(prioLower);
                              const isLåg = ["låg", "low", "may", "kan"].includes(prioLower);
                              
                              return (
                                <div key={rIdx} className="bg-white border border-slate-200 rounded-xl p-5 hover:border-slate-300 transition flex flex-col justify-between gap-4 shadow-sm">
                                  <div className="space-y-2">
                                    <div className="flex items-center justify-between gap-2 border-b border-slate-100 pb-2">
                                      <span className="text-xs font-bold leading-none font-mono text-indigo-600 bg-indigo-50 px-2.5 py-1 rounded border border-indigo-100 uppercase tracking-wider">
                                        {req.id}
                                      </span>
                                      
                                      <div className="flex gap-2">
                                        {req.verificationMethod && req.verificationMethod !== "Ej angivet" && (
                                          <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 border border-slate-200 px-2 py-0.5 rounded">
                                            {req.verificationMethod}
                                          </span>
                                        )}
                                        <span className={`text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded border ${
                                          isHög
                                            ? "text-rose-700 bg-rose-50 border-rose-200"
                                            : isLåg
                                            ? "text-slate-500 bg-slate-100 border-slate-200"
                                            : "text-amber-700 bg-amber-50 border-amber-200"
                                        }`}>
                                          {req.priority}
                                        </span>
                                      </div>
                                    </div>

                                    <h4 className="text-base font-bold text-slate-800 mt-1">{req.title}</h4>
                                    <p className="text-sm text-slate-600 font-light leading-relaxed font-sans mt-2">{req.description}</p>
                                  </div>

                                  <div className="border-t border-slate-100 pt-3 flex items-center justify-between text-xs text-slate-500 font-light mt-auto">
                                    <span>Kategori: <strong className="font-semibold text-slate-600">{req.category}</strong></span>
                                    {req.status && (
                                      <span className="font-mono bg-slate-50 px-1.5 py-0.5 rounded border border-slate-100">{req.status}</span>
                                    )}
                                  </div>
                                </div>
                              );
                            });
                          })()}
                        </div>

                      </div>
                    )}

                    {/* SUB-VIEW C: TECHNICAL PARAMETERS */}
                    {extractionType === "parameters" && (
                      <div className="bg-white border border-slate-200 rounded-2xl p-6 sm:p-8 space-y-6 shadow-sm">
                        <div className="flex items-center justify-between border-b border-slate-100 pb-4">
                          <div>
                            <span className="text-xs text-slate-400 font-bold block mb-1">DOKUMENTTITEL</span>
                            <h3 className="text-2xl font-extrabold text-slate-800 font-sans uppercase">
                              {(response.result as ParametersResult).documentTitle || "Parametrar"}
                            </h3>
                          </div>
                          <span className="text-xs font-mono text-slate-400">
                            {((response.result as ParametersResult).parameters || []).length} Parametrar extraherade
                          </span>
                        </div>

                        {/* Search and Table Grid */}
                        <div className="space-y-4">
                          <div className="relative max-w-sm">
                            <Search className="h-4 w-4 text-slate-400 absolute left-3 top-2.5" />
                            <input
                              type="text"
                              placeholder="Filtrera parametrar på namn..."
                              value={searchQuery}
                              onChange={(e) => setSearchQuery(e.target.value)}
                              className="bg-slate-50 border border-slate-200 text-xs text-slate-700 outline-none rounded-lg py-2 pl-9 pr-4 w-full focus:border-indigo-500 focus:ring-1 focus:ring-indigo-50 transition"
                            />
                          </div>

                          <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white shadow-sm">
                            <table className="w-full text-left border-collapse text-xs sm:text-sm font-sans">
                              <thead>
                                <tr className="bg-indigo-50/50 border-b border-indigo-100 text-indigo-900 font-bold">
                                  <th className="px-4 py-3">Parameternamn</th>
                                  <th className="px-4 py-3">Värde</th>
                                  <th className="px-4 py-3">Enhet</th>
                                  <th className="px-4 py-3">Tolerans / Gränsvärde</th>
                                  <th className="px-4 py-3">Sammanhang</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                {(() => {
                                  const paramList = (response.result as ParametersResult).parameters || [];
                                  const filtered = paramList.filter((p) =>
                                    p.name.toLowerCase().includes(searchQuery.toLowerCase()) || 
                                    (p.context && p.context.toLowerCase().includes(searchQuery.toLowerCase()))
                                  );

                                  if (filtered.length === 0) {
                                    return (
                                      <tr>
                                        <td colSpan={5} className="px-4 py-8 text-center text-slate-400 italic">
                                          Inga parametrar hittades för den valda filtreringen.
                                        </td>
                                      </tr>
                                    );
                                  }

                                  return filtered.map((param: ParameterItem, pIdx: number) => (
                                    <tr key={pIdx} className="hover:bg-slate-50/50 transition">
                                      <td className="px-4 py-3 text-slate-800 font-bold font-sans">{param.name}</td>
                                      <td className="px-4 py-3">
                                        <span className="font-mono font-semibold text-indigo-700 bg-indigo-50 px-2.5 py-1 rounded inline-block">
                                          {param.value}
                                        </span>
                                      </td>
                                      <td className="px-4 py-3 text-slate-500 font-mono">{param.unit || "-"}</td>
                                      <td className="px-4 py-3 text-amber-800 font-mono text-xs">{param.tolerance || "Börvärde"}</td>
                                      <td className="px-4 py-3 text-slate-500 font-light text-xs max-w-sm leading-relaxed">{param.context || "Ej angivet i dokumentet"}</td>
                                    </tr>
                                  ));
                                })()}
                              </tbody>
                            </table>
                          </div>
                        </div>

                      </div>
                    )}

                    {/* SUB-VIEW D: CUSTOM EXTRACTIONS */}
                    {extractionType === "custom" && (
                      <div className="bg-white border border-slate-200 rounded-2xl p-6 sm:p-8 space-y-6 shadow-sm">
                        <div className="border-b border-slate-100 pb-4">
                          <span className="text-xs text-slate-400 font-bold block mb-1">ANPASSAD SCHEMARESULTAT</span>
                          <h3 className="text-lg font-bold text-slate-800 mb-0 font-sans">
                            Anpassad datastruktur utvunnen via AI
                          </h3>
                        </div>

                        {/* Flex cards for dynamic keys in result */}
                        <div className="space-y-4">
                          <p className="text-xs text-slate-500 font-light">
                            Nedan visas den platta eller nästlade representationen av de unika datamängder du efterfrågade.
                          </p>

                          {/* Fallback code rendering of response.result with simple interactive visual cards */}
                          <div className="bg-slate-900 rounded-xl border border-slate-800 p-5 mt-4">
                            <pre className="text-xs font-mono text-emerald-400 whitespace-pre-wrap max-h-[500px] overflow-y-auto leading-relaxed">
                              {JSON.stringify(response.result, null, 2)}
                            </pre>
                          </div>
                        </div>
                      </div>
                    )}

                  </motion.div>
                )}

                {/* 2. CHROME-FORMATTED RAW JSON TAB */}
                {activeTab === "json" && (
                  <motion.div
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    className="space-y-4"
                  >
                    <div className="bg-[#0f172a] border border-slate-800 rounded-2xl relative overflow-hidden shadow-xl">
                      <div className="flex items-center justify-between px-4 py-3 bg-[#1e293b] border-b border-slate-800">
                        <div className="flex items-center gap-2">
                          <div className="flex gap-1.5">
                            <span className="w-2.5 h-2.5 rounded-full bg-rose-500 inline-block" />
                            <span className="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block" />
                            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block" />
                          </div>
                          <span className="text-xs font-mono text-slate-300 ml-2">output_{extractionType}.json</span>
                        </div>
                        <button
                          onClick={handleCopyJson}
                          className="h-8 border border-slate-750 hover:border-slate-605 bg-slate-900 hover:bg-slate-800 px-3 rounded-lg text-xs font-semibold text-slate-350 hover:text-white transition flex items-center gap-2 cursor-pointer"
                        >
                          {copied ? (
                            <>
                              <CheckCircle2 className="h-3.5 w-3.5 text-emerald-400" />
                              <span className="text-emerald-400">Kopierat!</span>
                            </>
                          ) : (
                            <>
                              <Copy className="h-3.5 w-3.5" />
                              <span>Kopiera</span>
                            </>
                          )}
                        </button>
                      </div>

                      <div className="p-4 sm:p-6 overflow-x-auto max-h-[600px] bg-[#0f172a] text-emerald-400">
                        <pre className="text-xs sm:text-sm font-mono leading-relaxed whitespace-pre font-normal select-text">
                          {JSON.stringify(response.result, null, 2)}
                        </pre>
                      </div>
                    </div>
                  </motion.div>
                )}

                {/* 3. RAW MARKDOWN PREVIEW TAB */}
                {activeTab === "raw_markdown" && (
                  <motion.div
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    className="space-y-4"
                  >
                    <div className="bg-white border border-slate-200 p-6 rounded-2xl shadow-sm space-y-4">
                      <p className="text-sm text-slate-600 leading-relaxed font-light">
                        Detta är den råa Markdown-texten som extraherades ur Word-dokumentets XML-strukturer innan den behandlades av Gemini. Mycket användbart för att granska innehållet utan Word.
                      </p>
                      <div className="bg-slate-50 rounded-xl border border-slate-200 p-5 max-h-[500px] overflow-y-auto">
                        <pre className="text-xs font-mono text-slate-700 leading-relaxed whitespace-pre-wrap select-text">
                          {response.rawMarkdown}
                        </pre>
                      </div>
                    </div>
                  </motion.div>
                )}

              </div>

              {/* Reset layout for new files */}
              <div className="flex justify-center pt-8 border-t border-slate-250">
                <button
                  onClick={handleReset}
                  className="px-6 py-3 border border-slate-200 hover:border-slate-300 bg-white hover:bg-slate-50 text-slate-700 hover:text-slate-900 rounded-xl transition text-sm font-bold flex items-center gap-2 cursor-pointer shadow-sm"
                >
                  <UploadCloud className="h-4 w-4 text-indigo-600" />
                  <span>Ladda upp ett till dokument</span>
                </button>
              </div>

            </motion.div>
          )}

        </AnimatePresence>

      </main>

      {/* Humble Footer */}
      <footer className="border-t border-slate-200 bg-white py-6 px-4 text-center text-xs text-slate-500 font-light mt-auto">
        <div className="max-w-6xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-4">
          <p>Mjukvara för Automationspartner AB • Word till JSON (Semantisk Schema-extraktion)</p>
          <p>© 2026 Word to JSON Converter • Genererad med precision</p>
        </div>
      </footer>

    </div>
  );
}
