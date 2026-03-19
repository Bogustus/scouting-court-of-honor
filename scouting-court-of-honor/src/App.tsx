import { useState, useCallback } from "react";
import { useDropzone } from "react-dropzone";
import { FileUp, File, X, Download, Mail } from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";
import { parseUploadFile, type ScoutAwardInfo } from "./lib/parser";
import { generateDocxScript, type ScriptOptions } from "./lib/generator";

export default function App() {
  const [fileRecords, setFileRecords] = useState<{ file: File; data: ScoutAwardInfo[] }[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);
  
  const data = fileRecords.flatMap(record => record.data);

  // Setup options
  const [options, setOptions] = useState<ScriptOptions>({
    date: new Date().toLocaleDateString("en-US", { month: "numeric", day: "numeric", year: "2-digit" }),
    time: "7:00PM-8:00PM",
    location: "St. Stephens",
    mc1Name: "Scout A",
    mc2Name: "Scout B",
    scoutmasterName: "Scoutmaster",
  });

  const onDrop = useCallback(async (acceptedFiles: File[]) => {
    const newRecords: { file: File; data: ScoutAwardInfo[] }[] = [];
    for (const file of acceptedFiles) {
      const parsed = await parseUploadFile(file);
      newRecords.push({ file, data: parsed });
    }
    setFileRecords((prev) => [...prev, ...newRecords]);
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({ onDrop, accept: {
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
    'application/vnd.ms-excel': ['.xls'],
    'text/csv': ['.csv'],
  } });

  const removeFile = (index: number) => {
    setFileRecords((prev) => prev.filter((_, i) => i !== index));
  };

  const handleGenerate = async () => {
    try {
      setIsGenerating(true);
      const blob = await generateDocxScript(data, options);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `CoH_Script_${options.date.replace(/\//g, "-")}.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      
      // Delay the revocation to ensure the browser has time to save the file
      setTimeout(() => {
        URL.revokeObjectURL(url);
      }, 2000);
    } catch (e) {
      console.error(e);
      alert("Failed to generate script. Please check the file formats.");
    } finally {
      setIsGenerating(false);
    }
  };

  const handleGenerateEmail = () => {
    if (data.length === 0) return;

    // Isolate unique scout names and format them
    const uniqueScouts = Array.from(new Set(data.map(d => d.originalName))).sort((a, b) => a.localeCompare(b));
    
    // Attempt to compute the day of the week
    let dayOfWeek = "Monday";
    try {
      const d = new Date(options.date);
      if (!isNaN(d.getTime())) {
        dayOfWeek = d.toLocaleDateString("en-US", { weekday: "long" });
      }
    } catch (e) {
      // ignore
    }

    // Clean up the time if it's "7:00PM-8:00PM" -> "7pm"
    let displayTime = options.time.split("-")[0];
    if (displayTime.includes(":00")) {
      displayTime = displayTime.replace(":00", "").toLowerCase();
    } else {
      displayTime = displayTime.toLowerCase();
    }

    const emailContent = `${dayOfWeek} night we have our next Court of Honor at ${displayTime} at ${options.location}.

The following scouts will be recognized at the Court of Honor:

${uniqueScouts.join('\n')}
`;

    const blob = new Blob([emailContent], { type: "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "CoH_Email_Draft.txt";
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  return (
    <div className="min-h-screen p-8 max-w-5xl mx-auto space-y-8">
      <header className="text-center space-y-4 pt-12">
        <motion.div
          initial={{ opacity: 0, scale: 0.9 }}
          animate={{ opacity: 1, scale: 1 }}
          className="inline-flex items-center justify-center p-4 bg-primary-600/20 rounded-full mb-4"
        >
          <FileUp size={48} className="text-primary-500" />
        </motion.div>
        <motion.h1 
          initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.1 }}
          className="text-4xl md:text-5xl font-extrabold tracking-tight"
        >
          Court of Honor <span className="text-transparent bg-clip-text bg-gradient-to-r from-primary-400 to-primary-600">Generator</span>
        </motion.h1>
        <motion.p
          initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.2 }}
          className="text-slate-400 text-lg max-w-2xl mx-auto"
        >
          Automate the tedious process of writing the Court of Honor script by dropping your Scoutbook / TroopWebHost exports right here.
        </motion.p>
      </header>

      <main className="grid md:grid-cols-2 gap-8">
        {/* Upload Column */}
        <motion.section 
          initial={{ opacity: 0, x: -20 }} animate={{ opacity: 1, x: 0 }} transition={{ delay: 0.3 }}
          className="bg-slate-800/50 backdrop-blur-xl border border-slate-700 p-6 rounded-3xl"
        >
          <h2 className="text-xl font-bold mb-4 flex items-center gap-2">
            <span className="bg-primary-500 text-white w-8 h-8 rounded-full flex items-center justify-center text-sm">1</span>
            Upload Data
          </h2>
          <div
            {...getRootProps()}
            className={`border-2 border-dashed rounded-2xl p-8 text-center cursor-pointer transition-all duration-300
              ${isDragActive ? "border-primary-500 bg-primary-500/10 scale-[1.02]" : "border-slate-600 hover:border-slate-500 bg-slate-800/50 hover:bg-slate-800"}
            `}
          >
            <input {...getInputProps()} />
            <FileUp size={40} className="mx-auto mb-4 text-slate-400" />
            <p className="text-slate-300 font-medium mb-1">
              {isDragActive ? "Drop files here..." : "Drag & drop your files"}
            </p>
            <p className="text-sm text-slate-500">Supports .xlsx, .xls, .csv</p>
          </div>

          <div className="mt-6 space-y-3">
            <AnimatePresence>
              {fileRecords.map((record, idx) => (
                <motion.div
                  key={idx + record.file.name}
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  className="flex items-center justify-between bg-slate-800 rounded-xl p-3 border border-slate-700"
                >
                  <div className="flex items-center gap-3 overflow-hidden">
                    <div className="bg-primary-500/20 p-2 rounded-lg text-primary-400">
                      <File size={20} />
                    </div>
                    <span className="truncate text-sm font-medium text-slate-300">{record.file.name}</span>
                  </div>
                  <button onClick={() => removeFile(idx)} className="text-slate-500 hover:text-red-400 transition-colors p-2">
                    <X size={18} />
                  </button>
                </motion.div>
              ))}
            </AnimatePresence>
            
            {fileRecords.length > 0 && (
              <p className="text-sm text-green-400 mt-4 flex items-center gap-2">
                <span className="relative flex h-3 w-3">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-green-400 opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-3 w-3 bg-green-500"></span>
                </span>
                Parsed {data.length} scout award records successfully.
              </p>
            )}
          </div>
        </motion.section>

        {/* Options Column */}
        <motion.section 
          initial={{ opacity: 0, x: 20 }} animate={{ opacity: 1, x: 0 }} transition={{ delay: 0.4 }}
          className="bg-slate-800/50 backdrop-blur-xl border border-slate-700 p-6 rounded-3xl"
        >
          <h2 className="text-xl font-bold mb-4 flex items-center gap-2">
            <span className="bg-primary-500 text-white w-8 h-8 rounded-full flex items-center justify-center text-sm">2</span>
            Configure Script
          </h2>
          
          <div className="space-y-4">
            <div className="grid grid-cols-2 gap-4">
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Date</span>
                <input
                  type="text"
                  value={options.date}
                  onChange={(e) => setOptions({ ...options, date: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all"
                />
              </label>
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Time</span>
                <input
                  type="text"
                  value={options.time}
                  onChange={(e) => setOptions({ ...options, time: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all"
                />
              </label>
            </div>
            
            <label className="block space-y-1.5">
              <span className="text-sm font-medium text-slate-400">Location</span>
              <input
                type="text"
                value={options.location}
                onChange={(e) => setOptions({ ...options, location: e.target.value })}
                className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all"
              />
            </label>

            <div className="grid grid-cols-2 gap-4 pt-2 border-t border-slate-700">
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">MC 1 Name</span>
                <input
                  type="text"
                  value={options.mc1Name}
                  onChange={(e) => setOptions({ ...options, mc1Name: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all"
                />
              </label>
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">MC 2 Name</span>
                <input
                  type="text"
                  value={options.mc2Name}
                  onChange={(e) => setOptions({ ...options, mc2Name: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all"
                />
              </label>
            </div>
            
            <label className="block space-y-1.5 pb-2">
              <span className="text-sm font-medium text-slate-400">Scoutmaster / Commissioner Name</span>
              <input
                type="text"
                value={options.scoutmasterName}
                onChange={(e) => setOptions({ ...options, scoutmasterName: e.target.value })}
                className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all"
              />
            </label>

            <div className="flex flex-col xl:flex-row gap-3 pt-2">
              <button
                onClick={handleGenerate}
                disabled={fileRecords.length === 0 || isGenerating}
                className={`flex-1 py-3 px-4 rounded-xl flex items-center justify-center gap-2 font-bold transition-all duration-300 ${
                  fileRecords.length > 0
                    ? "bg-primary-600 hover:bg-primary-500 text-white shadow-lg shadow-primary-500/25"
                    : "bg-slate-800 text-slate-500 cursor-not-allowed"
                }`}
              >
                <Download size={20} />
                {isGenerating ? "Generating..." : "Docx Script"}
              </button>

              <button
                onClick={handleGenerateEmail}
                disabled={fileRecords.length === 0 || isGenerating}
                className={`flex-1 py-3 px-4 rounded-xl flex items-center justify-center gap-2 font-bold transition-all duration-300 ${
                  fileRecords.length > 0
                    ? "bg-slate-700 hover:bg-slate-600 text-white shadow-lg shadow-slate-900/50"
                    : "bg-slate-800 text-slate-500 cursor-not-allowed"
                }`}
              >
                <Mail size={20} />
                Email Draft (.txt)
              </button>
            </div>
          </div>
        </motion.section>
      </main>
    </div>
  );
}
