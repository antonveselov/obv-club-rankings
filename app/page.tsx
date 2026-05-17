'use client';

import { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import Fuse from 'fuse.js';
import { 
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer
} from 'recharts';
import { X, TrendingUp, Calendar, Trophy, Medal, Info, MousePointer2, Search, User, Shield } from 'lucide-react';

interface Club {
  id: string;
  name: string;
}

interface Metadata {
  actualId: string;
  publications: { id: string; name: string }[];
  categories: { id: string; name: string }[];
  clubs: Club[];
}

interface Player {
  id: string;
  rank: string;
  name: string;
  club: string;
  points: string;
}

interface SearchResult {
  id: string;
  name: string;
  club: string;
  rankings: {
    singles?: number;
    doubles?: number;
    mixed?: number;
  };
}

interface CategoryResult {
  categoryName: string;
  players: Player[];
}

interface AttendanceRecord {
  name: string;
  count: number;
}

interface EligibilityResult {
  id: string;
  name: string;
  rank: string;
  clubRank: number;
  category: string;
  attendance: number;
  eligible: boolean;
}

interface RankHistoryPoint {
  date: string;
  rank: number | null;
}

interface PlayerHistory {
  singles: RankHistoryPoint[];
  doubles: RankHistoryPoint[];
  mixed: RankHistoryPoint[];
}

export default function Home() {
  const [metadata, setMetadata] = useState<Metadata | null>(null);
  const [mounted, setMounted] = useState(false);
  const [loadingMetadata, setLoadingMetadata] = useState(true);
  const [loadingPlayers, setLoadingPlayers] = useState(false);
  const [loadingTop5, setLoadingTop5] = useState(false);
  const [players, setPlayers] = useState<Player[]>([]);
  const [top5Results, setTop5Results] = useState<CategoryResult[]>([]);
  const [error, setError] = useState<string | null>(null);

  const [selectedClubId, setSelectedClubId] = useState('');
  const [selectedClubName, setSelectedClubName] = useState('');
  const [selectedPub, setSelectedPub] = useState('');
  const [selectedCat, setSelectedCat] = useState('');
  const [clubSearch, setClubSearch] = useState('');
  const [topX, setTopX] = useState(5);

  const [attendanceData, setAttendanceData] = useState<AttendanceRecord[]>([]);
  const [minTrainings, setMinTrainings] = useState(10);
  const [eligibilityResults, setEligibilityResults] = useState<EligibilityResult[]>([]);
  const [checkingEligibility, setCheckingEligibility] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [selectedPlayer, setSelectedPlayer] = useState<Player | null>(null);
  const [playerHistory, setPlayerHistory] = useState<PlayerHistory | null>(null);
  const [loadingHistory, setLoadingHistory] = useState(false);

  // Player Search State
  const [playerSearchQuery, setPlayerSearchQuery] = useState('');
  const [playerSearchResults, setPlayerSearchResults] = useState<SearchResult[]>([]);
  const [loadingSearch, setLoadingSearch] = useState(false);

  useEffect(() => {
    setMounted(true);
    fetch('/api/metadata')
      .then((res) => res.json())
      .then((data) => {
        if (data.error) throw new Error(data.error);
        setMetadata(data);
        if (data.publications.length > 0) setSelectedPub(data.publications[0].id);
        if (data.categories.length > 0) setSelectedCat(data.categories[0].id);
        setLoadingMetadata(false);
      })
      .catch((err) => {
        setError('Failed to load metadata: ' + err.message);
        setLoadingMetadata(false);
      });
  }, []);

  const handleSearch = async () => {
    if (!metadata || !selectedCat || !selectedClubId) return;
    setLoadingPlayers(true);
    setPlayers([]);
    setTop5Results([]);
    setEligibilityResults([]);
    setPlayerSearchResults([]);
    setError(null);

    try {
      const params = new URLSearchParams({
        actualId: metadata.actualId,
        catId: selectedCat,
        pubId: selectedPub,
        clubId: selectedClubId,
      });

      const res = await fetch(`/api/players?${params.toString()}`);
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setPlayers(data);
    } catch (err: any) {
      setError('Failed to load players: ' + err.message);
    } finally {
      setLoadingPlayers(false);
    }
  };

  const handleFetchTop5 = async () => {
    if (!metadata || !selectedClubId) return;
    setLoadingTop5(true);
    setTop5Results([]);
    setPlayers([]);
    setEligibilityResults([]);
    setPlayerSearchResults([]);
    setError(null);

    try {
      const promises = metadata.categories.map(async (cat) => {
        const params = new URLSearchParams({
          actualId: metadata.actualId,
          catId: cat.id,
          pubId: selectedPub,
          clubId: selectedClubId,
        });
        const res = await fetch(`/api/players?${params.toString()}`);
        const data = await res.json();
        if (!data.error) {
          return {
            categoryName: cat.name,
            players: data.slice(0, topX)
          };
        }
        return null;
      });

      const fetchedResults = await Promise.all(promises);
      setTop5Results(fetchedResults.filter((r): r is CategoryResult => r !== null));
    } catch (err: any) {
      setError('Failed to load Top 5: ' + err.message);
    } finally {
      setLoadingTop5(false);
    }
  };

  const handlePlayerSearch = async (e: React.FormEvent) => {
    e.preventDefault();
    if (playerSearchQuery.length < 3) return;
    
    setLoadingSearch(true);
    setPlayerSearchResults([]);
    setPlayers([]);
    setTop5Results([]);
    setEligibilityResults([]);
    setError(null);

    try {
      const res = await fetch(`/api/search-player?q=${encodeURIComponent(playerSearchQuery)}`);
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setPlayerSearchResults(data);
    } catch (err: any) {
      setError('Player Search Error: ' + err.message);
    } finally {
      setLoadingSearch(false);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[1] || wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];

        const records: AttendanceRecord[] = [];
        data.slice(1).forEach((row) => {
          const name = row[0];
          const count = parseInt(row[6]);
          if (name && !isNaN(count)) {
            records.push({ name: String(name).trim(), count });
          }
        });
        setAttendanceData(records);
        alert(`Erfolgreich ${records.length} Teilnehmer aus ${wsname} geladen.`);
      } catch (err) {
        setError('Fehler beim Lesen der Excel-Datei.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const checkEligibility = async () => {
    if (!metadata || !selectedClubId || attendanceData.length === 0) {
      setError('Bitte zuerst Club auswählen und Anwesenheitsliste hochladen.');
      return;
    }

    setCheckingEligibility(true);
    setPlayers([]);
    setTop5Results([]);
    setPlayerSearchResults([]);
    setError(null);

    try {
      const allEligible: EligibilityResult[] = [];
      const fuse = new Fuse(attendanceData, { keys: ['name'], threshold: 0.3 });

      for (const cat of metadata.categories) {
        const params = new URLSearchParams({
          actualId: metadata.actualId,
          catId: cat.id,
          pubId: selectedPub,
          clubId: selectedClubId,
        });
        const res = await fetch(`/api/players?${params.toString()}`);
        const data = await res.json();
        
        if (data.error) continue;

        const topPlayers = data.slice(0, topX);
        topPlayers.forEach((p: Player, idx: number) => {
          const match = fuse.search(p.name);
          if (match.length > 0) {
            const attendance = match[0].item.count;
            if (attendance >= minTrainings) {
              allEligible.push({
                id: p.id,
                name: p.name,
                rank: p.rank,
                clubRank: idx + 1,
                category: cat.name,
                attendance: attendance,
                eligible: true
              });
            }
          }
        });
      }
      setEligibilityResults(allEligible);
    } catch (err: any) {
      setError('Fehler beim Prüfen der Berechtigung: ' + err.message);
    } finally {
      setCheckingEligibility(false);
    }
  };

  const openPlayerHistory = async (player: { id: string; name: string }) => {
    if (!metadata || !player.id) {
      console.warn("Cannot fetch history: Missing player ID");
      return;
    }
    setSelectedPlayer({ id: player.id, name: player.name, rank: '', club: '', points: '' });
    setPlayerHistory(null);
    setLoadingHistory(true);
    
    try {
      const params = new URLSearchParams({
        actualId: metadata.actualId,
        playerId: player.id,
        clubId: selectedClubId || ''
      });
      const res = await fetch(`/api/history?${params.toString()}`);
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setPlayerHistory(data);
    } catch (err: any) {
      console.error(err);
    } finally {
      setLoadingHistory(false);
    }
  };

  if (loadingMetadata) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <p className="text-gray-600 font-bold uppercase tracking-widest text-xs">Initialisierung...</p>
        </div>
      </div>
    );
  }

  const filteredClubs = metadata?.clubs.filter(c => 
    c.name.toLowerCase().includes(clubSearch.toLowerCase())
  ) || [];

  const CustomTooltip = ({ active, payload, label }: any) => {
    if (active && payload && payload.length) {
      return (
        <div className="bg-white p-3 border border-gray-200 shadow-2xl rounded-xl">
          <p className="text-[10px] text-gray-400 mb-1 font-black uppercase tracking-widest">{label}</p>
          <p className="text-lg font-black text-gray-900 leading-tight">
            Rang <span className="text-blue-600">#{payload[0].value}</span>
          </p>
        </div>
      );
    }
    return null;
  };

  return (
    <main className="min-h-screen bg-gray-50 py-8 px-4 sm:px-6 lg:px-8 font-sans">
      <div className="max-w-4xl mx-auto">
        
        {/* Global Player Search */}
        <div className="mb-8 relative">
            <form onSubmit={handlePlayerSearch} className="group relative">
                <input
                    type="text"
                    className="w-full bg-white border-2 border-gray-100 rounded-3xl p-6 pl-14 shadow-xl focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 outline-none transition-all font-black text-lg placeholder:text-gray-300"
                    placeholder="Globaler Spieler-Quick-Search (z.B. Anton Veselov)..."
                    value={playerSearchQuery}
                    onChange={(e) => setPlayerSearchQuery(e.target.value)}
                />
                <Search className="absolute left-6 top-1/2 -translate-y-1/2 h-6 w-6 text-gray-300 group-focus-within:text-blue-500 transition-colors" />
                <button 
                    type="submit"
                    disabled={playerSearchQuery.length < 3 || loadingSearch}
                    className="absolute right-3 top-1/2 -translate-y-1/2 bg-gray-900 text-white px-6 py-3 rounded-2xl font-black text-xs uppercase tracking-widest hover:bg-black transition-all active:scale-95 disabled:bg-gray-100 disabled:text-gray-300"
                >
                    {loadingSearch ? '...' : 'SUCHEN'}
                </button>
            </form>
        </div>

        {playerSearchResults.length > 0 && (
          <div className="bg-white shadow-2xl rounded-3xl overflow-hidden border border-blue-100 mb-8 animate-in slide-in-from-top-4 duration-500">
            <div className="px-8 py-5 bg-blue-600 flex justify-between items-center">
                <h3 className="text-sm font-black text-white uppercase tracking-widest flex items-center gap-2">
                    <User className="h-4 w-4" /> Suchergebnisse
                </h3>
                <button onClick={() => setPlayerSearchResults([])} className="text-white/60 hover:text-white">
                    <X className="h-5 w-5" />
                </button>
            </div>
            <div className="divide-y divide-gray-50">
              {playerSearchResults.map((res) => (
                <div key={res.id} className="p-6 hover:bg-blue-50/30 transition-all group cursor-pointer flex justify-between items-center" onClick={() => openPlayerHistory(res)}>
                    <div className="flex items-center gap-6">
                        <div className="bg-gray-100 p-3 rounded-2xl group-hover:bg-blue-100 transition-colors text-gray-400 group-hover:text-blue-600">
                            <User className="h-6 w-6" />
                        </div>
                        <div>
                            <p className="text-xl font-black text-gray-900 group-hover:text-blue-600 transition-colors">{res.name}</p>
                            <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mt-1 flex items-center gap-2">
                                <Shield className="h-3 w-3" /> {res.club}
                            </p>
                        </div>
                    </div>
                    <div className="flex gap-3">
                        {res.rankings.singles && (
                            <div className="text-center px-4 py-2 bg-gray-50 rounded-xl border border-gray-100">
                                <p className="text-[8px] font-black text-gray-400 uppercase tracking-tighter">Singles</p>
                                <p className="text-sm font-black text-gray-900">#{res.rankings.singles}</p>
                            </div>
                        )}
                        {res.rankings.doubles && (
                            <div className="text-center px-4 py-2 bg-gray-50 rounded-xl border border-gray-100">
                                <p className="text-[8px] font-black text-gray-400 uppercase tracking-tighter">Doubles</p>
                                <p className="text-sm font-black text-gray-900">#{res.rankings.doubles}</p>
                            </div>
                        )}
                        {res.rankings.mixed && (
                            <div className="text-center px-4 py-2 bg-gray-50 rounded-xl border border-gray-100">
                                <p className="text-[8px] font-black text-gray-400 uppercase tracking-tighter">Mixed</p>
                                <p className="text-sm font-black text-gray-900">#{res.rankings.mixed}</p>
                            </div>
                        )}
                        <div className="self-center ml-4 opacity-0 group-hover:opacity-100 transition-opacity">
                            <TrendingUp className="h-6 w-6 text-blue-500" />
                        </div>
                    </div>
                </div>
              ))}
            </div>
          </div>
        )}

        <div className="bg-white shadow-sm rounded-3xl p-6 sm:p-8 mb-8 border border-gray-200">
          <h1 className="text-2xl font-black text-gray-900 mb-8 flex items-center gap-4 tracking-tighter">
            <div className="bg-yellow-400 p-2 rounded-xl shadow-inner">
                <Trophy className="text-white h-6 w-6" />
            </div>
            ÖBV CLUB EXPLORER
          </h1>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            <div className="relative">
              <label className="block text-xs font-black text-gray-400 uppercase tracking-widest mb-2 ml-1">Club auswählen</label>
              <input
                type="text"
                className="block w-full border-2 border-gray-100 rounded-2xl shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm border p-4 bg-gray-50 transition-all font-bold placeholder:text-gray-300"
                placeholder="Club-Name..."
                value={clubSearch}
                onChange={(e) => {
                  setClubSearch(e.target.value);
                  setSelectedClubId('');
                }}
              />
              {clubSearch && !selectedClubId && (
                <div className="absolute z-30 mt-2 w-full bg-white shadow-2xl max-h-60 rounded-2xl py-2 text-base ring-1 ring-black ring-opacity-5 overflow-auto focus:outline-none sm:text-sm border border-gray-100 backdrop-blur-xl">
                  {filteredClubs.length > 0 ? (
                    filteredClubs.map((club) => (
                      <div
                        key={club.id}
                        className="cursor-pointer select-none relative py-3 pl-5 pr-9 hover:bg-blue-600 hover:text-white text-gray-800 transition-all font-bold"
                        onClick={() => {
                          setSelectedClubId(club.id);
                          setSelectedClubName(club.name);
                          setClubSearch(club.name);
                        }}
                      >
                        {club.name}
                      </div>
                    ))
                  ) : (
                    <div className="py-3 pl-5 pr-9 text-gray-400 italic">Kein Verein gefunden</div>
                  )}
                </div>
              )}
            </div>

            <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-black text-gray-400 uppercase tracking-widest mb-2 ml-1">Kategorie</label>
                  <select
                    className="block w-full border-2 border-gray-100 rounded-2xl shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm border p-4 bg-gray-50 font-bold transition-all"
                    value={selectedCat}
                    onChange={(e) => setSelectedCat(e.target.value)}
                  >
                    {metadata?.categories.map((cat) => (
                      <option key={cat.id} value={cat.id}>{cat.name}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-black text-gray-400 uppercase tracking-widest mb-2 ml-1 flex items-center gap-1">
                    Datum
                  </label>
                  <select
                    className="block w-full border-2 border-gray-100 rounded-2xl shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm border p-4 bg-gray-50 font-bold transition-all"
                    value={selectedPub}
                    onChange={(e) => setSelectedPub(e.target.value)}
                  >
                    {metadata?.publications.map((pub) => (
                      <option key={pub.id} value={pub.id}>{pub.name}</option>
                    ))}
                  </select>
                </div>
            </div>

            <div className="md:col-span-2 flex flex-col sm:flex-row gap-4">
              <button
                onClick={handleSearch}
                disabled={loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId}
                className={`flex-1 flex justify-center py-4 px-6 border border-transparent rounded-2xl shadow-lg text-sm font-black text-white transition-all uppercase tracking-widest ${
                  loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId ? 'bg-gray-200 text-gray-400' : 'bg-gray-900 hover:bg-black active:scale-95'
                } focus:outline-none h-14 items-center`}
              >
                {loadingPlayers ? 'Lade...' : 'Einzelne Liste'}
              </button>
              
              <div className="flex flex-1 gap-4">
                <div className="w-24">
                  <input
                    type="number"
                    min="1"
                    max="50"
                    className="block w-full border-2 border-gray-100 rounded-2xl shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm border p-4 h-14 text-center bg-gray-50 font-black text-lg"
                    value={topX}
                    onChange={(e) => setTopX(parseInt(e.target.value) || 1)}
                  />
                </div>
                <button
                  onClick={handleFetchTop5}
                  disabled={loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId}
                  className={`flex-grow flex justify-center py-4 px-6 border-2 border-blue-600 rounded-2xl shadow-sm text-sm font-black transition-all uppercase tracking-widest ${
                    loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId ? 'text-gray-300 border-gray-200' : 'text-blue-600 hover:bg-blue-50 active:scale-95'
                  } focus:outline-none h-14 items-center`}
                >
                  {loadingTop5 ? `Lade Top-${topX}...` : `Top-${topX} Übersicht`}
                </button>
              </div>
            </div>
          </div>

          <div className="mt-12 pt-8 border-t border-gray-100">
            <div className="flex items-center gap-3 mb-6">
                <div className="bg-green-100 p-2 rounded-xl">
                    <Medal className="text-green-600 h-6 w-6" />
                </div>
                <h2 className="text-xl font-black text-gray-900 uppercase tracking-tight">
                  TURNIER-RÜCKERSTATTUNG
                </h2>
            </div>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8 items-end">
              <div>
                <label className="block text-xs font-black text-gray-400 uppercase tracking-widest mb-2 ml-1">Mind. Trainings (K)</label>
                <input
                  type="number"
                  min="0"
                  className="block w-full border-2 border-gray-100 rounded-2xl shadow-sm focus:ring-green-500 focus:border-green-500 sm:text-sm border p-4 bg-gray-50 font-black text-lg"
                  value={minTrainings}
                  onChange={(e) => setMinTrainings(parseInt(e.target.value) || 0)}
                />
              </div>
              <div>
                <label className="block text-xs font-black text-gray-400 uppercase tracking-widest mb-2 ml-1">Anwesenheitsliste (.xlsx)</label>
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  ref={fileInputRef}
                  onChange={handleFileUpload}
                  className="block w-full text-sm text-gray-400 file:mr-4 file:py-3 file:px-6 file:rounded-xl file:border-0 file:text-xs file:font-black file:bg-green-600 file:text-white hover:file:bg-green-700 cursor-pointer border-2 border-gray-100 rounded-2xl bg-gray-50 h-14"
                />
              </div>
              <div className="md:col-span-2">
                <button
                  onClick={checkEligibility}
                  disabled={loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId || attendanceData.length === 0}
                  className={`w-full flex justify-center py-4 px-6 border border-transparent rounded-2xl shadow-xl text-base font-black text-white transition-all uppercase tracking-widest ${
                    loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId || attendanceData.length === 0 ? 'bg-gray-200' : 'bg-green-600 hover:bg-green-700 active:scale-95 shadow-green-200/50'
                  } focus:outline-none h-16 items-center`}
                >
                  {checkingEligibility ? 'PRÜFE...' : 'ELIGIBILITY CHECK STARTEN'}
                </button>
              </div>
            </div>
          </div>

          {error && (
            <div className="mt-6 p-5 bg-red-50 border-2 border-red-100 text-red-700 rounded-2xl text-sm font-bold flex items-center gap-3 animate-pulse">
              <Info className="h-6 w-6 shrink-0" /> {error}
            </div>
          )}
        </div>

        {eligibilityResults.length > 0 && (
          <div className="bg-white shadow-2xl rounded-3xl overflow-hidden border-2 border-green-500/20 mb-12 animate-in fade-in zoom-in duration-500">
            <div className="px-8 py-6 bg-green-600 flex justify-between items-center">
              <div>
                <h3 className="text-2xl leading-6 font-black text-white uppercase tracking-tight">
                    Berechtigte Spieler
                </h3>
                <p className="text-sm text-green-100 font-bold mt-1 uppercase tracking-widest">Top-{topX} im Club & mind. {minTrainings} Trainings</p>
              </div>
              <Medal className="h-10 w-10 text-white/50" />
            </div>
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-gray-100">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-8 py-5 text-left text-[10px] font-black text-gray-400 uppercase tracking-[0.2em]">Name</th>
                    <th className="px-8 py-5 text-left text-[10px] font-black text-gray-400 uppercase tracking-[0.2em]">Kategorie</th>
                    <th className="px-8 py-5 text-center text-[10px] font-black text-gray-400 uppercase tracking-[0.2em]">Club-Rang</th>
                    <th className="px-8 py-5 text-center text-[10px] font-black text-gray-400 uppercase tracking-[0.2em]">ÖBV-Rang</th>
                    <th className="px-8 py-5 text-right text-[10px] font-black text-gray-400 uppercase tracking-[0.2em]">Trainings</th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-50">
                  {eligibilityResults.map((res, idx) => (
                    <tr key={idx} className="hover:bg-green-50 transition-colors group">
                      <td className="px-8 py-5 whitespace-nowrap text-base font-black text-gray-900 cursor-pointer group-hover:text-blue-600 flex items-center gap-2" onClick={() => openPlayerHistory({ id: res.id, name: res.name })}>
                        {res.name} <MousePointer2 className="h-3 w-3 opacity-0 group-hover:opacity-100" />
                      </td>
                      <td className="px-8 py-5 whitespace-nowrap text-sm text-gray-500 font-bold">{res.category}</td>
                      <td className="px-8 py-5 whitespace-nowrap text-sm text-center font-black text-blue-600 bg-blue-50/30">{res.clubRank}.</td>
                      <td className="px-8 py-5 whitespace-nowrap text-sm text-center text-gray-300 font-black tracking-tighter italic">#{res.rank}</td>
                      <td className="px-8 py-5 whitespace-nowrap text-sm text-right font-black text-green-600 bg-green-50/50">{res.attendance}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {players.length > 0 && (
          <div className="bg-white shadow-2xl rounded-3xl overflow-hidden border border-gray-100 mt-8 animate-in slide-in-from-bottom-8 duration-700">
            <div className="px-8 py-6 bg-gray-900 flex justify-between items-center">
              <h3 className="text-xl leading-6 font-black text-white uppercase tracking-widest">
                {selectedClubName}
              </h3>
              <div className="bg-blue-600 px-3 py-1 rounded-full text-[10px] font-black text-white">{players.length} SPIELER</div>
            </div>
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-gray-100">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-8 py-5 text-left text-[10px] font-black text-gray-400 uppercase tracking-[0.2em] w-24">ÖBV</th>
                    <th className="px-8 py-5 text-left text-[10px] font-black text-gray-400 uppercase tracking-[0.2em] w-24">Club</th>
                    <th className="px-8 py-5 text-left text-[10px] font-black text-gray-400 uppercase tracking-[0.2em]">Spieler</th>
                    <th className="px-8 py-5 text-left text-[10px] font-black text-gray-400 uppercase tracking-[0.2em]">Verein</th>
                    <th className="px-8 py-5 text-right text-[10px] font-black text-gray-400 uppercase tracking-[0.2em]">Punkte</th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-50">
                  {players.map((player, idx) => (
                    <tr key={idx} className="hover:bg-blue-50/50 transition-colors group">
                      <td className="px-8 py-5 whitespace-nowrap text-sm font-black text-gray-300 italic">#{player.rank}</td>
                      <td className="px-8 py-5 whitespace-nowrap text-sm text-blue-600 font-black">{idx + 1}.</td>
                      <td className="px-8 py-5 whitespace-nowrap text-base font-black text-gray-900 cursor-pointer group-hover:text-blue-700 flex items-center gap-2" onClick={() => openPlayerHistory(player)}>
                        {player.name} <MousePointer2 className="h-3 w-3 opacity-0 group-hover:opacity-100 text-blue-400" />
                      </td>
                      <td className="px-8 py-5 whitespace-nowrap text-sm text-gray-400 font-bold">{player.club}</td>
                      <td className="px-8 py-5 whitespace-nowrap text-sm text-gray-900 text-right font-black tabular-nums">{player.points}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {top5Results.length > 0 && (
          <div className="mt-12 space-y-10 animate-in fade-in duration-1000">
            <div className="flex items-center gap-6">
                <h2 className="text-3xl font-black text-gray-900 uppercase tracking-tighter italic shrink-0">TOP-{topX} SNAPSHOT</h2>
                <div className="h-[2px] flex-grow bg-gradient-to-r from-gray-200 to-transparent rounded-full"></div>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8 sm:gap-10">
              {top5Results.map((result, idx) => (
                <div key={idx} className="bg-white shadow-xl rounded-[2rem] border border-gray-100 overflow-hidden hover:shadow-2xl transition-all duration-300 group">
                  <div className="px-6 py-5 bg-gray-50 border-b border-gray-50 flex justify-between items-center group-hover:bg-blue-600 transition-colors">
                    <h3 className="text-[10px] font-black text-gray-400 uppercase tracking-[0.2em] group-hover:text-white transition-colors">{result.categoryName}</h3>
                    <TrendingUp className="h-4 w-4 text-gray-300 group-hover:text-white transition-colors" />
                  </div>
                  <div className="p-4 sm:p-5">
                    {result.players.length > 0 ? (
                      <ul className="space-y-2">
                        {result.players.map((p, pIdx) => (
                          <li key={pIdx} className="px-5 py-4 flex justify-between items-center rounded-2xl hover:bg-blue-50 transition-all group/item cursor-pointer border border-transparent hover:border-blue-100" onClick={() => openPlayerHistory(p)}>
                            <div className="flex items-center gap-5">
                                <span className="text-xl font-black text-blue-600 w-8">{pIdx + 1}.</span>
                                <div>
                                    <p className="text-base font-black text-gray-900 group-hover/item:text-blue-700 transition-colors leading-none mb-1.5">{p.name}</p>
                                    <p className="text-[10px] text-gray-300 font-black uppercase tracking-widest">ÖBV RANKING #{p.rank}</p>
                                </div>
                            </div>
                            <span className="text-[11px] font-black text-gray-400 bg-gray-50 px-3 py-1.5 rounded-lg tabular-nums border border-gray-100">{p.points}</span>
                          </li>
                        ))}
                      </ul>
                    ) : (
                      <p className="p-10 text-xs text-gray-300 uppercase tracking-widest text-center font-black">No Players</p>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {!loadingPlayers && !loadingTop5 && !checkingEligibility && !loadingSearch && playerSearchResults.length === 0 && selectedClubId && players.length === 0 && top5Results.length === 0 && eligibilityResults.length === 0 && !error && (
          <div className="text-center py-32 bg-white rounded-[3rem] border-4 border-dashed border-gray-50 shadow-inner mt-8 group">
            <div className="bg-gray-50 w-20 h-20 rounded-full flex items-center justify-center mx-auto mb-6 group-hover:scale-110 transition-transform">
                <TrendingUp className="h-10 w-10 text-gray-200" />
            </div>
            <p className="text-gray-300 font-black uppercase tracking-[0.3em] text-xs">Wählen Sie eine Aktion aus</p>
          </div>
        )}
      </div>

      {/* History Modal */}
      {selectedPlayer && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6" aria-labelledby="modal-title" role="dialog" aria-modal="true">
          <div className="fixed inset-0 bg-gray-950/80 backdrop-blur-md transition-opacity" aria-hidden="true" onClick={() => setSelectedPlayer(null)}></div>
          
          <div className="relative bg-white rounded-[2.5rem] text-left overflow-hidden shadow-2xl transform transition-all w-full max-w-4xl max-h-[92vh] flex flex-col border border-white/20">
            <div className="p-8 sm:p-10 flex-shrink-0 border-b border-gray-50 flex justify-between items-start bg-white">
                <div>
                  <div className="flex items-center gap-3 mb-2">
                    <div className="h-2 w-2 rounded-full bg-blue-600 animate-pulse"></div>
                    <span className="text-[10px] font-black text-blue-600 uppercase tracking-[0.2em]">Live Player Analysis</span>
                  </div>
                  <h3 className="text-4xl font-black text-gray-900 tracking-tighter uppercase leading-none" id="modal-title">
                    {selectedPlayer.name}
                  </h3>
                  <div className="text-gray-400 font-black uppercase tracking-[0.2em] text-xs mt-3 flex items-center gap-2">
                    <div className="h-4 w-[1px] bg-gray-200"></div>
                    {selectedClubName || 'Österreichischer Verband'}
                  </div>
                </div>
                <button onClick={() => setSelectedPlayer(null)} className="p-3 hover:bg-gray-100 rounded-2xl transition-all active:scale-90 bg-gray-50">
                  <X className="h-8 w-8 text-gray-900" />
                </button>
            </div>

            <div className="flex-grow overflow-y-auto p-8 sm:p-10 bg-white">
                {loadingHistory ? (
                  <div className="py-40 text-center">
                    <div className="relative h-24 w-24 mx-auto mb-10">
                        <div className="absolute inset-0 border-8 border-gray-50 rounded-full"></div>
                        <div className="absolute inset-0 border-8 border-t-blue-600 rounded-full animate-spin"></div>
                        <div className="absolute inset-4 bg-blue-50 rounded-full flex items-center justify-center">
                            <TrendingUp className="h-6 w-6 text-blue-600" />
                        </div>
                    </div>
                    <p className="text-gray-900 font-black text-2xl tracking-tighter uppercase">Generiere Report...</p>
                    <p className="text-gray-400 font-bold mt-3 uppercase tracking-widest text-[10px]">Analyse der letzten 12 Monate läuft</p>
                  </div>
                ) : playerHistory && mounted ? (
                  <div className="space-y-20">
                    {/* Singles Chart */}
                    <div className="bg-white rounded-[2rem] p-4 sm:p-2">
                      <div className="flex justify-between items-end mb-10 px-4">
                        <div>
                            <h4 className="text-xs font-black text-gray-300 uppercase tracking-[0.3em] mb-2">
                                Discipline 01
                            </h4>
                            <p className="text-xl font-black text-gray-900 tracking-tight uppercase">Einzel-Verlauf</p>
                        </div>
                        <div className="px-4 py-2 bg-blue-50 text-blue-600 rounded-xl text-[10px] font-black tracking-widest border border-blue-100">YEARLY TREND</div>
                      </div>
                      <div className="h-72 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <AreaChart data={playerHistory.singles} margin={{ top: 0, right: 10, left: -20, bottom: 0 }}>
                            <defs>
                              <linearGradient id="colorRank" x1="0" y1="0" x2="0" y2="1">
                                <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.3}/>
                                <stop offset="95%" stopColor="#3b82f6" stopOpacity={0}/>
                              </linearGradient>
                            </defs>
                            <CartesianGrid strokeDasharray="6 6" vertical={false} stroke="#f1f5f9" />
                            <XAxis 
                                dataKey="date" 
                                tick={{fontSize: 9, fill: '#cbd5e1', fontWeight: 900}} 
                                tickLine={false}
                                axisLine={false}
                                dy={15}
                            />
                            <YAxis 
                                reversed={false}
                                domain={['auto', 'auto']} 
                                tick={{fontSize: 10, fill: '#cbd5e1', fontWeight: 900}} 
                                tickLine={false}
                                axisLine={false}
                            />
                            <Tooltip content={<CustomTooltip />} />
                            <Area type="monotone" dataKey="rank" stroke="#3b82f6" strokeWidth={5} fillOpacity={1} fill="url(#colorRank)" />
                          </AreaChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    {/* Doubles Chart */}
                    <div className="bg-white rounded-[2rem] p-4 sm:p-2">
                      <div className="flex justify-between items-end mb-10 px-4">
                        <div>
                            <h4 className="text-xs font-black text-gray-300 uppercase tracking-[0.3em] mb-2">
                                Discipline 02
                            </h4>
                            <p className="text-xl font-black text-gray-900 tracking-tight uppercase">Doppel-Verlauf</p>
                        </div>
                        <div className="px-4 py-2 bg-green-50 text-green-600 rounded-xl text-[10px] font-black tracking-widest border border-green-100">YEARLY TREND</div>
                      </div>
                      <div className="h-72 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <AreaChart data={playerHistory.doubles} margin={{ top: 0, right: 10, left: -20, bottom: 0 }}>
                            <defs>
                              <linearGradient id="colorRankGreen" x1="0" y1="0" x2="0" y2="1">
                                <stop offset="5%" stopColor="#10b981" stopOpacity={0.3}/>
                                <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                              </linearGradient>
                            </defs>
                            <CartesianGrid strokeDasharray="6 6" vertical={false} stroke="#f1f5f9" />
                            <XAxis dataKey="date" tick={{fontSize: 9, fill: '#cbd5e1', fontWeight: 900}} tickLine={false} axisLine={false} dy={15} />
                            <YAxis reversed={false} domain={['auto', 'auto']} tick={{fontSize: 10, fill: '#cbd5e1', fontWeight: 900}} tickLine={false} axisLine={false} />
                            <Tooltip content={<CustomTooltip />} />
                            <Area type="monotone" dataKey="rank" stroke="#10b981" strokeWidth={5} fillOpacity={1} fill="url(#colorRankGreen)" />
                          </AreaChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    {/* Mixed Chart */}
                    <div className="bg-white rounded-[2rem] p-4 sm:p-2 mb-10">
                      <div className="flex justify-between items-end mb-10 px-4">
                        <div>
                            <h4 className="text-xs font-black text-gray-300 uppercase tracking-[0.3em] mb-2">
                                Discipline 03
                            </h4>
                            <p className="text-xl font-black text-gray-900 tracking-tight uppercase">Mixed-Verlauf</p>
                        </div>
                        <div className="px-4 py-2 bg-purple-50 text-purple-600 rounded-xl text-[10px] font-black tracking-widest border border-purple-100">YEARLY TREND</div>
                      </div>
                      <div className="h-72 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <AreaChart data={playerHistory.mixed} margin={{ top: 0, right: 10, left: -20, bottom: 0 }}>
                            <defs>
                              <linearGradient id="colorRankPurple" x1="0" y1="0" x2="0" y2="1">
                                <stop offset="5%" stopColor="#a855f7" stopOpacity={0.3}/>
                                <stop offset="95%" stopColor="#a855f7" stopOpacity={0}/>
                              </linearGradient>
                            </defs>
                            <CartesianGrid strokeDasharray="6 6" vertical={false} stroke="#f1f5f9" />
                            <XAxis dataKey="date" tick={{fontSize: 9, fill: '#cbd5e1', fontWeight: 900}} tickLine={false} axisLine={false} dy={15} />
                            <YAxis reversed={false} domain={['auto', 'auto']} tick={{fontSize: 10, fill: '#cbd5e1', fontWeight: 900}} tickLine={false} axisLine={false} />
                            <Tooltip content={<CustomTooltip />} />
                            <Area type="monotone" dataKey="rank" stroke="#a855f7" strokeWidth={5} fillOpacity={1} fill="url(#colorRankPurple)" />
                          </AreaChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="py-20 text-center text-red-500 font-bold uppercase tracking-[0.2em]">
                    Data Unavailable for this player.
                  </div>
                )}
            </div>
            
            <div className="p-8 sm:p-10 border-t border-gray-50 flex-shrink-0 bg-white">
                <button type="button" onClick={() => setSelectedPlayer(null)} className="w-full py-5 px-8 bg-gray-900 rounded-[1.5rem] text-sm font-black text-white hover:bg-black transition-all active:scale-95 shadow-xl shadow-gray-200 uppercase tracking-widest">
                  Back to Dashboard
                </button>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
