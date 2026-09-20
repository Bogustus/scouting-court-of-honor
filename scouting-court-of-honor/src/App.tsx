import { useState, useCallback, useEffect } from "react";
import { useDropzone } from "react-dropzone";
import { FileUp, File, X, Download, Mail } from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";
import { parseUploadFile, type ScoutAwardInfo } from "./lib/parser";
import { generateDocxScript, type ScriptOptions } from "./lib/generator";
import { compareByLastName } from "./lib/sort";

function getNextMondayDate() {
  const d = new Date();
  d.setDate(d.getDate() + ((1 + 7 - d.getDay()) % 7 || 7));
  return d.toLocaleDateString("en-US", { month: "numeric", day: "numeric", year: "2-digit" });
}

function extractDateFromFilename(filename: string): string | null {
  const match = filename.match(/(\d{1,2})[._-](\d{1,2})(?:[._-](\d{2,4}))?/);
  if (match) {
    const month = match[1];
    const day = match[2];
    let yearRaw = match[3];
    
    if (!yearRaw) {
      const yearMatch = filename.match(/(20\d{2})/);
      if (yearMatch) {
         yearRaw = yearMatch[1];
      }
    }

    let year = new Date().getFullYear();
    if (yearRaw) {
       year = yearRaw.length === 2 ? parseInt(`20${yearRaw}`) : parseInt(yearRaw);
    }
    
    const d = new Date(year, parseInt(month) - 1, parseInt(day));
    if (!isNaN(d.getTime())) {
      return d.toLocaleDateString("en-US", { month: "numeric", day: "numeric", year: "2-digit" });
    }
  }
  return null;
}

export interface ExtendedOptions extends ScriptOptions {
  advChairsName: string;
  advChairsEmail: string;
  leadershipName: string;
  leadershipEmail: string;
  bruceName: string;
  bruceEmail: string;
  youthLeadersName: string;
  youthLeadersEmail: string;
  troopListEmail: string;
  mcEmails: string;
  colorGuardEmails: string;
  googleDocLink: string;
}

export default function App() {
  const [fileRecords, setFileRecords] = useState<{ file: File; data: ScoutAwardInfo[] }[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);
  
  const data = fileRecords.flatMap(record => record.data);

  // Setup options with LocalStorage Persistence
  const [options, setOptions] = useState<ExtendedOptions>(() => {
    const saved = localStorage.getItem("coh_options");
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        // Ensure color guard exists in case migrating from old state
        return { colorGuardNames: "Scout C, Scout D", introTitle: "Unit Commissioner", ...parsed };
      } catch (e) {}
    }
    return {
      date: getNextMondayDate(),
      time: "7:00PM-8:00PM",
      location: "St. Stephens",
      mc1Name: "Scout A",
      mc2Name: "Scout B",
      colorGuardNames: "Scout C, Scout D",
      scoutmasterName: "Bruce McGurk",
      introTitle: "Unit Commissioner",
      advChairsName: "Denise and Katherine",
      advChairsEmail: "",
      leadershipName: "Michael and Dennis",
      leadershipEmail: "",
      bruceName: "Bruce",
      bruceEmail: "",
      youthLeadersName: "Ira and Jordan",
      youthLeadersEmail: "",
      troopListEmail: "troop303-orinda@googlegroups.com",
      mcEmails: "",
      colorGuardEmails: "",
      googleDocLink: "https://docs.google.com/document/d/..."
    };
  });

  useEffect(() => {
    localStorage.setItem("coh_options", JSON.stringify(options));
  }, [options]);

  const onDrop = useCallback(async (acceptedFiles: File[]) => {
    const newRecords: { file: File; data: ScoutAwardInfo[] }[] = [];
    let extractedDate: string | null = null;
    
    for (const file of acceptedFiles) {
      const parsed = await parseUploadFile(file);
      newRecords.push({ file, data: parsed });
      
      // Attempt to extract date from the first file we see
      if (!extractedDate) {
        extractedDate = extractDateFromFilename(file.name);
      }
    }
    setFileRecords((prev) => [...prev, ...newRecords]);

    // Update the options explicitly if a date was found in the filename
    if (extractedDate) {
      setOptions(prev => ({ ...prev, date: extractedDate as string }));
    }
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
      let fileName = `CoH Script.docx`;
      try {
        const d = new Date(options.date);
        if (!isNaN(d.getTime())) {
          const year = d.getFullYear();
          const month = d.getMonth() + 1;
          const day = d.getDate();
          fileName = `${year} ${month}-${day} CoH Script.docx`;
        }
      } catch (e) {}
      a.download = fileName;
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

  const buildEmailLink = (to: string, subject: string, body: string) => {
    return `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(to)}&su=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  };

  const getUniqueScouts = () => Array.from(new Set(data.map(d => d.scoutName))).sort(compareByLastName);

  const handleGenerateEmail = () => {
    if (data.length === 0) return;
    const uniqueScouts = getUniqueScouts();
    
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
            
            <div className="grid grid-cols-2 gap-4 pt-2 border-t border-slate-700">
              <label className="block space-y-1.5 pb-2">
                <span className="text-sm font-medium text-slate-400">Intro Speaker Name</span>
                <input
                  type="text"
                  value={options.scoutmasterName}
                  onChange={(e) => setOptions({ ...options, scoutmasterName: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all"
                />
              </label>
              
              <label className="block space-y-1.5 pb-2">
                <span className="text-sm font-medium text-slate-400">Intro Speaker Title</span>
                <input
                  type="text"
                  value={options.introTitle}
                  onChange={(e) => setOptions({ ...options, introTitle: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all"
                />
              </label>
            </div>

            <div className="pt-2 border-t border-slate-700">
              <label className="block space-y-1.5 pb-2">
                <span className="text-sm font-medium text-slate-400">Color Guard Names</span>
                <input
                  type="text"
                  value={options.colorGuardNames}
                  onChange={(e) => setOptions({ ...options, colorGuardNames: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all"
                />
              </label>
            </div>

            <hr className="border-slate-700 my-2" />
            <h3 className="text-lg font-bold">Coordination Contacts (Names vs. Emails)</h3>
            
            <div className="grid grid-cols-2 gap-4">
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Adv. Chairs Name(s)</span>
                <input
                  type="text"
                  value={options.advChairsName}
                  onChange={(e) => setOptions({ ...options, advChairsName: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs"
                />
              </label>
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Adv. Chairs Email(s)</span>
                <input
                  type="text"
                  value={options.advChairsEmail}
                  onChange={(e) => setOptions({ ...options, advChairsEmail: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs placeholder:text-slate-600"
                  placeholder="denise@..., kath@..."
                />
              </label>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Bruce's Name</span>
                <input
                  type="text"
                  value={options.bruceName}
                  onChange={(e) => setOptions({ ...options, bruceName: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs"
                />
              </label>
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Bruce's Email</span>
                <input
                  type="text"
                  value={options.bruceEmail}
                  onChange={(e) => setOptions({ ...options, bruceEmail: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs placeholder:text-slate-600"
                  placeholder="bruce@..."
                />
              </label>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Leadership Team Name(s)</span>
                <input
                  type="text"
                  value={options.leadershipName}
                  onChange={(e) => setOptions({ ...options, leadershipName: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs"
                />
              </label>
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Leadership Team Email(s)</span>
                <input
                  type="text"
                  value={options.leadershipEmail}
                  onChange={(e) => setOptions({ ...options, leadershipEmail: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs"
                />
              </label>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Youth Leaders (SPLs)</span>
                <input
                  type="text"
                  value={options.youthLeadersName}
                  onChange={(e) => setOptions({ ...options, youthLeadersName: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs"
                />
              </label>
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Youth Leaders Email(s)</span>
                <input
                  type="text"
                  value={options.youthLeadersEmail}
                  onChange={(e) => setOptions({ ...options, youthLeadersEmail: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs"
                />
              </label>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">MC Email(s)</span>
                <input
                  type="text"
                  value={options.mcEmails}
                  onChange={(e) => setOptions({ ...options, mcEmails: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs placeholder:text-slate-600"
                  placeholder="For milestone 6 emailing"
                />
              </label>
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Color Guard Email(s)</span>
                <input
                  type="text"
                  value={options.colorGuardEmails}
                  onChange={(e) => setOptions({ ...options, colorGuardEmails: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs placeholder:text-slate-600"
                />
              </label>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Troop Distro Email</span>
                <input
                  type="text"
                  value={options.troopListEmail}
                  onChange={(e) => setOptions({ ...options, troopListEmail: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs"
                />
              </label>
              <label className="block space-y-1.5">
                <span className="text-sm font-medium text-slate-400">Google Doc Script Link</span>
                <input
                  type="text"
                  value={options.googleDocLink}
                  onChange={(e) => setOptions({ ...options, googleDocLink: e.target.value })}
                  className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2 text-slate-200 focus:outline-none focus:ring-2 focus:ring-primary-500 transition-all text-xs placeholder:text-slate-600"
                  placeholder="https://docs.google.com/..."
                />
              </label>
            </div>

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
        <motion.section 
          initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.5 }}
          className="col-span-1 md:col-span-2 bg-slate-800/50 backdrop-blur-xl border border-slate-700 p-6 rounded-3xl mt-4"
        >
          <h2 className="text-xl font-bold mb-6 flex items-center gap-2">
            <span className="bg-primary-500 text-white w-8 h-8 rounded-full flex items-center justify-center text-sm">3</span>
            Milestone Tracker / Automations
          </h2>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <a 
              href={buildEmailLink(
                options.advChairsEmail, 
                `Advancement and Court of Honor ${options.date}`,
                `Hey ${options.advChairsName}, we have a Court of Honor on Monday, ${options.date}. Where are we on the data for that? If I can get the report by Saturday morning, I can make the script for the MC's so they have time to look at it over the weekend.`
              )}
              target="_blank" rel="noopener noreferrer"
              className="p-4 bg-slate-800 rounded-xl border border-slate-700 hover:border-primary-500 transition-all flex items-start gap-3 group"
            >
              <div className="bg-slate-700 p-2 rounded-lg text-slate-300 group-hover:bg-primary-500 group-hover:text-white transition-colors">
                <Mail size={20} />
              </div>
              <div>
                <h3 className="font-bold text-slate-200">1. Data Request (3 Weeks Out)</h3>
                <p className="text-xs text-slate-400 mt-1">To: Advancement Chairs</p>
              </div>
            </a>

            <a 
              href={buildEmailLink(
                options.bruceEmail, 
                `Troop 303 Court of Honor - ${options.date}`,
                `Hey ${options.bruceName}, will you be able to make it to the Court of Honor on Monday, ${options.date}? It's always nice to have you there to officially open and close the ceremony.`
              )}
              target="_blank" rel="noopener noreferrer"
              className="p-4 bg-slate-800 rounded-xl border border-slate-700 hover:border-primary-500 transition-all flex items-start gap-3 group"
            >
              <div className="bg-slate-700 p-2 rounded-lg text-slate-300 group-hover:bg-primary-500 group-hover:text-white transition-colors">
                <Mail size={20} />
              </div>
              <div>
                <h3 className="font-bold text-slate-200">2. Bruce's Reminder</h3>
                <p className="text-xs text-slate-400 mt-1">To: Bruce (No CCs)</p>
              </div>
            </a>

            <a 
              href={buildEmailLink(
                [options.youthLeadersEmail, options.leadershipEmail].filter(Boolean).join(", "), 
                `MC & Color Guard Recruitment - ${options.date}`,
                `Hi ${options.youthLeadersName},\n\nI’m starting to build out the team for our upcoming CoH on Monday, ${options.date}. Who will be serving as our MCs and who is assigned to the Color Guard?\n\nI’d like to get their names into the script as soon as possible. As a reminder, they need to arrive at 6:00 PM for rehearsal.`
              )}
              target="_blank" rel="noopener noreferrer"
              className="p-4 bg-slate-800 rounded-xl border border-slate-700 hover:border-primary-500 transition-all flex items-start gap-3 group"
            >
              <div className="bg-slate-700 p-2 rounded-lg text-slate-300 group-hover:bg-primary-500 group-hover:text-white transition-colors">
                <Mail size={20} />
              </div>
              <div>
                <h3 className="font-bold text-slate-200">3. Recruitment Request</h3>
                <p className="text-xs text-slate-400 mt-1">To: SPLs, Leadership Team</p>
              </div>
            </a>

            <a 
              href={buildEmailLink(
                options.troopListEmail, 
                `Recognition Preview - ${options.date}`,
                `The following scouts will be recognized at the Court of Honor on ${options.date}:\n\n${getUniqueScouts().map(s => `• ${s}`).join('\n')}\n\nPlease let me know if any corrections are needed.`
              )}
              target="_blank" rel="noopener noreferrer"
              onClick={(e) => { if (data.length === 0) { e.preventDefault(); alert("Upload a spreadsheet to populate scout names first!"); } }}
              className={`p-4 bg-slate-800 rounded-xl border ${data.length > 0 ? "border-slate-700 hover:border-primary-500 cursor-pointer group" : "border-red-900/50 opacity-50 cursor-not-allowed"} transition-all flex items-start gap-3`}
            >
              <div className={`p-2 rounded-lg transition-colors ${data.length > 0 ? "bg-slate-700 text-slate-300 group-hover:bg-primary-500 group-hover:text-white" : "bg-slate-800 text-slate-600"}`}>
                <Mail size={20} />
              </div>
              <div>
                <h3 className="font-bold text-slate-200">4. Parent Recognition Preview</h3>
                <p className="text-xs text-slate-400 mt-1">To: Troop Distro (Requires Parsed File)</p>
              </div>
            </a>

            <a 
              href={buildEmailLink(
                options.advChairsEmail, 
                `Script Ready for Review - ${options.date}`,
                `The script is ready. Here is the link to the google doc: ${options.googleDocLink}\n\nPlease correct any mistakes you see before I send it out to the scouts.`
              )}
              target="_blank" rel="noopener noreferrer"
              className="p-4 bg-slate-800 rounded-xl border border-slate-700 hover:border-primary-500 transition-all flex items-start gap-3 group"
            >
              <div className="bg-slate-700 p-2 rounded-lg text-slate-300 group-hover:bg-primary-500 group-hover:text-white transition-colors">
                <Mail size={20} />
              </div>
              <div>
                <h3 className="font-bold text-slate-200">5. Coordinator Review</h3>
                <p className="text-xs text-slate-400 mt-1">To: Advancement Chairs</p>
              </div>
            </a>

            <a 
              href={buildEmailLink(
                [options.mcEmails, options.colorGuardEmails, options.youthLeadersEmail].filter(Boolean).join(", "), 
                `Final Script & Rehearsal Instructions - ${options.date}`,
                `${options.mc1Name} and ${options.mc2Name}, thank-you for being our MC's. ${options.colorGuardNames}, thank-you for being the Color Guard.\n\nPlease arrive at 6:00 PM (one hour early) for rehearsal. Here is the link to the script: ${options.googleDocLink}\n\nMCs, please be prepared to pronounce all scout names correctly.`
              )}
              target="_blank" rel="noopener noreferrer"
              className="p-4 bg-slate-800 rounded-xl border border-slate-700 hover:border-primary-500 transition-all flex items-start gap-3 group"
            >
              <div className="bg-slate-700 p-2 rounded-lg text-slate-300 group-hover:bg-primary-500 group-hover:text-white transition-colors">
                <Mail size={20} />
              </div>
              <div>
                <h3 className="font-bold text-slate-200">6. Final Script & Logistics</h3>
                <p className="text-xs text-slate-400 mt-1">To: MCs, Color Guard, SPLs</p>
              </div>
            </a>

          </div>
        </motion.section>
      </main>
    </div>
  );
}
