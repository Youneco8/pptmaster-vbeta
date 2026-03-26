import React, { useState, useRef } from 'react';
import { GoogleGenAI } from "@google/genai";
import * as XLSX from 'xlsx';
import pptxgen from "pptxgenjs";
import { 
  Upload, 
  FileText, 
  Image as ImageIcon, 
  ChevronRight, 
  Download, 
  Loader2, 
  CheckCircle2, 
  AlertCircle,
  Presentation,
  RefreshCw,
  Search,
  MessageSquare
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// Utility for tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface SlideData {
  number: number;
  originalTitle: string;
  actionTitle?: string;
  visual?: string; // base64
  visualMimeType?: string;
  visualAnalysis?: string;
  additionalData?: string;
  status: 'pending' | 'processing' | 'completed' | 'error';
}

const MCKINSEY_BLUE = "#002D72";

export default function App() {
  const [slides, setSlides] = useState<SlideData[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [currentStep, setCurrentStep] = useState<'upload' | 'processing' | 'review'>('upload');
  const [error, setError] = useState<string | null>(null);
  const [progress, setProgress] = useState({ current: 0, total: 0 });

  const excelInputRef = useRef<HTMLInputElement>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);

  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet);

      // Expecting columns: "slide number" (or similar) and "title"
      const newSlides: SlideData[] = jsonData.map((row: any) => {
        const slideNum = row['slide number'] || row['Slide Number'] || row['col1'] || Object.values(row)[0];
        const title = row['title'] || row['Title'] || row['col2'] || Object.values(row)[1];
        
        return {
          number: Number(slideNum),
          originalTitle: String(title),
          status: 'pending' as const
        };
      }).filter(s => !isNaN(s.number));

      setSlides(newSlides.sort((a, b) => a.number - b.number));
      setError(null);
    } catch (err) {
      setError("Failed to parse Excel file. Please ensure it has two columns: Slide Number and Title.");
    }
  };

  const handleImageUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    const newSlides = [...slides];
    
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const fileName = file.name.toUpperCase();
      // Match S1.png, S01.png, Slide1.png etc.
      const match = fileName.match(/S(?:LIDE)?\s*(\d+)/);
      
      if (match) {
        const slideNum = parseInt(match[1]);
        const slideIndex = newSlides.findIndex(s => s.number === slideNum);
        
        if (slideIndex !== -1) {
          const base64 = await fileToBase64(file);
          newSlides[slideIndex] = {
            ...newSlides[slideIndex],
            visual: base64,
            visualMimeType: file.type
          };
        }
      }
    }
    
    setSlides(newSlides);
  };

  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        const result = reader.result as string;
        resolve(result.split(',')[1]); // Remove data:image/png;base64,
      };
      reader.onerror = error => reject(error);
    });
  };

  const processWithAI = async () => {
    if (slides.length === 0) {
      setError("Please upload an Excel file first.");
      return;
    }

    setIsProcessing(true);
    setCurrentStep('processing');
    setProgress({ current: 0, total: slides.length });

    const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY! });
    const model = "gemini-3.1-pro-preview";

    const updatedSlides = [...slides];

    for (let i = 0; i < updatedSlides.length; i++) {
      const slide = updatedSlides[i];
      updatedSlides[i].status = 'processing';
      setSlides([...updatedSlides]);

      try {
        const parts: any[] = [
          { text: `Act as a Senior Consultant at a top-tier firm (McKinsey, BCG, Bain). 
          I am providing you with a slide title and potentially a visual.
          
          TASK:
          1. Transform the title "${slide.originalTitle}" into a punchy, executive-level "action title" (a full sentence that conveys the key takeaway).
          2. If a visual is provided, analyze it and provide 2-3 bullet points of high-level business commentary/insights.
          3. Search for additional analytical data, market trends, or supporting facts related to this topic ("${slide.originalTitle}") to enrich the slide.
          
          OUTPUT FORMAT:
          Return ONLY a JSON object with these keys:
          {
            "actionTitle": "string",
            "visualAnalysis": "string (bullet points)",
            "additionalData": "string (bullet points)"
          }` }
        ];

        if (slide.visual) {
          parts.push({
            inlineData: {
              data: slide.visual,
              mimeType: slide.visualMimeType || "image/png"
            }
          });
        }

        const result = await ai.models.generateContent({
          model,
          contents: [{ parts }],
          config: {
            responseMimeType: "application/json",
            tools: [{ googleSearch: {} }]
          }
        });

        const response = JSON.parse(result.text || "{}");
        
        updatedSlides[i] = {
          ...slide,
          actionTitle: response.actionTitle,
          visualAnalysis: response.visualAnalysis,
          additionalData: response.additionalData,
          status: 'completed'
        };
      } catch (err) {
        console.error(`Error processing slide ${slide.number}:`, err);
        updatedSlides[i].status = 'error';
      }

      setProgress(prev => ({ ...prev, current: i + 1 }));
      setSlides([...updatedSlides]);
    }

    setIsProcessing(false);
    setCurrentStep('review');
  };

  const generatePPTX = () => {
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";

    slides.forEach(slide => {
      const pptSlide = pres.addSlide();
      
      // Background / Theme
      pptSlide.background = { color: "FFFFFF" };

      // Action Title
      pptSlide.addText(slide.actionTitle || slide.originalTitle, {
        x: 0.5,
        y: 0.3,
        w: 9.0,
        h: 0.8,
        fontSize: 24,
        bold: true,
        color: MCKINSEY_BLUE,
        fontFace: "Arial"
      });

      // Divider line
      pptSlide.addShape(pres.ShapeType.line, {
        x: 0.5,
        y: 1.1,
        w: 9.0,
        h: 0,
        line: { color: MCKINSEY_BLUE, width: 1 }
      });

      if (slide.visual) {
        // Visual on the left
        pptSlide.addImage({
          data: `data:${slide.visualMimeType};base64,${slide.visual}`,
          x: 0.5,
          y: 1.3,
          w: 5.5,
          h: 3.5
        });

        // Commentary on the right
        pptSlide.addText("Executive Insights", {
          x: 6.2,
          y: 1.3,
          w: 3.3,
          h: 0.3,
          fontSize: 12,
          bold: true,
          color: MCKINSEY_BLUE
        });

        pptSlide.addText(slide.visualAnalysis || "No visual analysis available.", {
          x: 6.2,
          y: 1.7,
          w: 3.3,
          h: 1.5,
          fontSize: 10,
          color: "333333",
          bullet: true
        });

        // Additional Data below commentary
        pptSlide.addText("Market Context & Data", {
          x: 6.2,
          y: 3.3,
          w: 3.3,
          h: 0.3,
          fontSize: 12,
          bold: true,
          color: MCKINSEY_BLUE
        });

        pptSlide.addText(slide.additionalData || "No additional data found.", {
          x: 6.2,
          y: 3.7,
          w: 3.3,
          h: 1.5,
          fontSize: 10,
          color: "333333",
          bullet: true
        });
      } else {
        // Full width text if no visual
        pptSlide.addText("Strategic Overview", {
          x: 0.5,
          y: 1.3,
          w: 9.0,
          h: 0.3,
          fontSize: 14,
          bold: true,
          color: MCKINSEY_BLUE
        });

        pptSlide.addText(slide.additionalData || "No additional data found.", {
          x: 0.5,
          y: 1.7,
          w: 9.0,
          h: 3.5,
          fontSize: 12,
          color: "333333",
          bullet: true
        });
      }

      // Footer
      pptSlide.addText(`Slide ${slide.number} | Confidential`, {
        x: 0.5,
        y: 5.3,
        w: 9.0,
        h: 0.3,
        fontSize: 8,
        color: "999999",
        align: "right"
      });
    });

    pres.writeFile({ fileName: `Executive_Deck_${new Date().getTime()}.pptx` });
  };

  return (
    <div className="min-h-screen flex flex-col">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-8 py-4 flex items-center justify-between sticky top-0 z-10">
        <div className="flex items-center gap-3">
          <div className="bg-[#002D72] p-2 rounded">
            <Presentation className="text-white w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight text-[#002D72]">Executive Deck Builder</h1>
            <p className="text-xs text-gray-500 font-mono uppercase tracking-widest">Consulting AI Assistant</p>
          </div>
        </div>
        
        <div className="flex items-center gap-4">
          {currentStep === 'review' && (
            <button
              onClick={generatePPTX}
              className="flex items-center gap-2 bg-[#002D72] text-white px-4 py-2 rounded-md hover:bg-[#001D4A] transition-colors font-medium text-sm shadow-sm"
            >
              <Download className="w-4 h-4" />
              Download PPTX
            </button>
          )}
          {currentStep !== 'upload' && !isProcessing && (
            <button
              onClick={() => {
                setSlides([]);
                setCurrentStep('upload');
              }}
              className="text-gray-500 hover:text-gray-700 text-sm font-medium flex items-center gap-1"
            >
              <RefreshCw className="w-4 h-4" />
              Reset
            </button>
          )}
        </div>
      </header>

      <main className="flex-1 max-w-6xl mx-auto w-full p-8">
        <AnimatePresence mode="wait">
          {currentStep === 'upload' && (
            <motion.div
              key="upload"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="grid grid-cols-1 md:grid-cols-2 gap-8"
            >
              {/* Excel Upload Card */}
              <div className="consulting-card p-8 rounded-xl flex flex-col items-center text-center group cursor-pointer hover:border-[#002D72] transition-all"
                   onClick={() => excelInputRef.current?.click()}>
                <input 
                  type="file" 
                  ref={excelInputRef} 
                  className="hidden" 
                  accept=".xlsx, .xls" 
                  onChange={handleExcelUpload}
                />
                <div className="w-16 h-16 bg-blue-50 rounded-full flex items-center justify-center mb-6 group-hover:bg-blue-100 transition-colors">
                  <FileText className="w-8 h-8 text-[#002D72]" />
                </div>
                <h3 className="text-lg font-semibold mb-2">1. Upload Slide Outline</h3>
                <p className="text-sm text-gray-500 mb-6">
                  Upload an Excel file with slide numbers and titles. 
                  Gemini will transform these into action titles.
                </p>
                {slides.length > 0 ? (
                  <div className="flex items-center gap-2 text-green-600 bg-green-50 px-3 py-1 rounded-full text-xs font-medium">
                    <CheckCircle2 className="w-3 h-3" />
                    {slides.length} Slides Loaded
                  </div>
                ) : (
                  <div className="text-xs font-mono text-gray-400 uppercase tracking-wider">
                    Click to browse files
                  </div>
                )}
              </div>

              {/* Image Upload Card */}
              <div className={cn(
                "consulting-card p-8 rounded-xl flex flex-col items-center text-center group transition-all",
                slides.length === 0 ? "opacity-50 cursor-not-allowed" : "cursor-pointer hover:border-[#002D72]"
              )}
                   onClick={() => slides.length > 0 && imageInputRef.current?.click()}>
                <input 
                  type="file" 
                  ref={imageInputRef} 
                  className="hidden" 
                  multiple 
                  accept="image/png" 
                  onChange={handleImageUpload}
                />
                <div className="w-16 h-16 bg-blue-50 rounded-full flex items-center justify-center mb-6 group-hover:bg-blue-100 transition-colors">
                  <ImageIcon className="w-8 h-8 text-[#002D72]" />
                </div>
                <h3 className="text-lg font-semibold mb-2">2. Upload Visuals (PNG)</h3>
                <p className="text-sm text-gray-500 mb-6">
                  Name files as S1.png, S2.png, etc. to match slide numbers.
                  Gemini will analyze graphs and charts.
                </p>
                {slides.some(s => s.visual) ? (
                  <div className="flex items-center gap-2 text-green-600 bg-green-50 px-3 py-1 rounded-full text-xs font-medium">
                    <CheckCircle2 className="w-3 h-3" />
                    {slides.filter(s => s.visual).length} Visuals Mapped
                  </div>
                ) : (
                  <div className="text-xs font-mono text-gray-400 uppercase tracking-wider">
                    {slides.length === 0 ? "Upload Excel first" : "Click to browse PNGs"}
                  </div>
                )}
              </div>

              {/* Action Button */}
              <div className="md:col-span-2 flex flex-col items-center gap-4 mt-8">
                {error && (
                  <div className="flex items-center gap-2 text-red-600 bg-red-50 px-4 py-2 rounded-lg text-sm">
                    <AlertCircle className="w-4 h-4" />
                    {error}
                  </div>
                )}
                <button
                  disabled={slides.length === 0 || isProcessing}
                  onClick={processWithAI}
                  className={cn(
                    "flex items-center gap-3 px-8 py-4 rounded-lg font-bold text-lg transition-all shadow-lg",
                    slides.length > 0 
                      ? "bg-[#002D72] text-white hover:bg-[#001D4A] scale-105" 
                      : "bg-gray-200 text-gray-400 cursor-not-allowed"
                  )}
                >
                  Start AI Transformation
                  <ChevronRight className="w-5 h-5" />
                </button>
              </div>
            </motion.div>
          )}

          {currentStep === 'processing' && (
            <motion.div
              key="processing"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex flex-col items-center justify-center py-20"
            >
              <div className="relative mb-8">
                <Loader2 className="w-20 h-20 text-[#002D72] animate-spin" />
                <div className="absolute inset-0 flex items-center justify-center font-bold text-[#002D72]">
                  {Math.round((progress.current / progress.total) * 100)}%
                </div>
              </div>
              <h2 className="text-2xl font-bold mb-2">Generating Strategic Insights</h2>
              <p className="text-gray-500 mb-8">Acting as Senior Consultant to refine titles and analyze visuals...</p>
              
              <div className="w-full max-w-md bg-gray-200 h-2 rounded-full overflow-hidden">
                <motion.div 
                  className="bg-[#002D72] h-full"
                  initial={{ width: 0 }}
                  animate={{ width: `${(progress.current / progress.total) * 100}%` }}
                />
              </div>
              <p className="mt-4 text-sm font-mono text-gray-400">Processing Slide {progress.current} of {progress.total}</p>
            </motion.div>
          )}

          {currentStep === 'review' && (
            <motion.div
              key="review"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="space-y-8"
            >
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-2xl font-bold">Review Generated Deck</h2>
                <div className="text-sm text-gray-500">
                  {slides.length} slides ready for export
                </div>
              </div>

              <div className="grid grid-cols-1 gap-6">
                {slides.map((slide) => (
                  <div key={slide.number} className="consulting-card rounded-xl overflow-hidden flex flex-col md:flex-row">
                    {/* Slide Preview Left */}
                    <div className="w-full md:w-1/3 bg-gray-50 p-6 border-r border-gray-100 flex flex-col">
                      <div className="flex items-center justify-between mb-4">
                        <span className="text-xs font-mono font-bold text-[#002D72] bg-blue-50 px-2 py-1 rounded">SLIDE {slide.number}</span>
                        {slide.status === 'completed' ? (
                          <CheckCircle2 className="w-4 h-4 text-green-500" />
                        ) : (
                          <AlertCircle className="w-4 h-4 text-red-500" />
                        )}
                      </div>
                      
                      {slide.visual ? (
                        <div className="aspect-video bg-white border border-gray-200 rounded overflow-hidden mb-4">
                          <img 
                            src={`data:${slide.visualMimeType};base64,${slide.visual}`} 
                            alt={`Slide ${slide.number}`}
                            className="w-full h-full object-contain"
                            referrerPolicy="no-referrer"
                          />
                        </div>
                      ) : (
                        <div className="aspect-video bg-white border border-dashed border-gray-300 rounded flex items-center justify-center mb-4 text-gray-400 text-xs italic">
                          No visual uploaded
                        </div>
                      )}
                      
                      <div className="mt-auto">
                        <p className="text-[10px] font-mono text-gray-400 uppercase tracking-tighter mb-1">Original Title</p>
                        <p className="text-xs text-gray-600 italic">"{slide.originalTitle}"</p>
                      </div>
                    </div>

                    {/* AI Insights Right */}
                    <div className="flex-1 p-8 space-y-6">
                      <div>
                        <div className="flex items-center gap-2 mb-2">
                          <FileText className="w-4 h-4 text-[#002D72]" />
                          <h4 className="text-xs font-bold uppercase tracking-widest text-gray-400">Action Title</h4>
                        </div>
                        <p className="text-lg font-semibold text-[#002D72] leading-tight">
                          {slide.actionTitle || slide.originalTitle}
                        </p>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <div>
                          <div className="flex items-center gap-2 mb-2">
                            <Search className="w-4 h-4 text-[#002D72]" />
                            <h4 className="text-xs font-bold uppercase tracking-widest text-gray-400">Visual Insights</h4>
                          </div>
                          <div className="text-sm text-gray-700 whitespace-pre-line prose prose-sm max-w-none">
                            {slide.visualAnalysis || "No visual analysis generated."}
                          </div>
                        </div>
                        
                        <div>
                          <div className="flex items-center gap-2 mb-2">
                            <MessageSquare className="w-4 h-4 text-[#002D72]" />
                            <h4 className="text-xs font-bold uppercase tracking-widest text-gray-400">Market Context</h4>
                          </div>
                          <div className="text-sm text-gray-700 whitespace-pre-line prose prose-sm max-w-none">
                            {slide.additionalData || "No additional data found."}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Footer */}
      <footer className="bg-white border-t border-gray-200 px-8 py-4 text-center text-[10px] text-gray-400 uppercase tracking-[0.2em] font-mono">
        Executive Deck Builder &copy; 2026 | Powered by Gemini 3.1 Pro
      </footer>
    </div>
  );
}
