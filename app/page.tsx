'use client';

import { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import Fuse from 'fuse.js';

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
  rank: string;
  name: string;
  club: string;
  points: string;
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
  name: string;
  rank: string;
  clubRank: number;
  category: string;
  attendance: number;
  eligible: boolean;
}

export default function Home() {
  const [metadata, setMetadata] = useState<Metadata | null>(null);
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

  // Reimbursement State
  const [attendanceData, setAttendanceData] = useState<AttendanceRecord[]>([]);
  const [minTrainings, setMinTrainings] = useState(10);
  const [eligibilityResults, setEligibilityResults] = useState<EligibilityResult[]>([]);
  const [checkingEligibility, setCheckingEligibility] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
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
          const count = parseInt(row[6]); // Column G is index 6
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

        // Take top X from this category
        const topPlayers = data.slice(0, topX);
        topPlayers.forEach((p: Player, idx: number) => {
          const match = fuse.search(p.name);
          if (match.length > 0) {
            const attendance = match[0].item.count;
            if (attendance >= minTrainings) {
              allEligible.push({
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

  if (loadingMetadata) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <p className="text-gray-600">Lade Vereinsdaten...</p>
        </div>
      </div>
    );
  }

  const filteredClubs = metadata?.clubs.filter(c => 
    c.name.toLowerCase().includes(clubSearch.toLowerCase())
  ) || [];

  return (
    <main className="min-h-screen bg-gray-50 py-8 px-4 sm:px-6 lg:px-8">
      <div className="max-w-4xl mx-auto">
        <div className="bg-white shadow-sm rounded-lg p-6 mb-8 border border-gray-200">
          <h1 className="text-2xl font-bold text-gray-900 mb-6">ÖBV Club Spieler Ranking & Rückerstattung</h1>
          
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="relative">
              <label className="block text-sm font-medium text-gray-700 mb-1">Verein auswählen</label>
              <input
                type="text"
                className="block w-full border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm border p-2"
                placeholder="Suche Verein..."
                value={clubSearch}
                onChange={(e) => {
                  setClubSearch(e.target.value);
                  setSelectedClubId('');
                }}
              />
              {clubSearch && !selectedClubId && (
                <div className="absolute z-10 mt-1 w-full bg-white shadow-lg max-h-60 rounded-md py-1 text-base ring-1 ring-black ring-opacity-5 overflow-auto focus:outline-none sm:text-sm">
                  {filteredClubs.length > 0 ? (
                    filteredClubs.map((club) => (
                      <div
                        key={club.id}
                        className="cursor-pointer select-none relative py-2 pl-3 pr-9 hover:bg-blue-600 hover:text-white text-gray-900"
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
                    <div className="py-2 pl-3 pr-9 text-gray-500">Kein Verein gefunden</div>
                  )}
                </div>
              )}
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Kategorie</label>
              <select
                className="block w-full border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm border p-2"
                value={selectedCat}
                onChange={(e) => setSelectedCat(e.target.value)}
              >
                {metadata?.categories.map((cat) => (
                  <option key={cat.id} value={cat.id}>{cat.name}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Datum (RL-Woche)</label>
              <select
                className="block w-full border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm border p-2"
                value={selectedPub}
                onChange={(e) => setSelectedPub(e.target.value)}
              >
                {metadata?.publications.map((pub) => (
                  <option key={pub.id} value={pub.id}>{pub.name}</option>
                ))}
              </select>
            </div>

            <div className="flex flex-col gap-2">
              <button
                onClick={handleSearch}
                disabled={loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId}
                className={`w-full flex justify-center py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white ${
                  loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId ? 'bg-gray-400' : 'bg-blue-600 hover:bg-blue-700'
                } focus:outline-none h-10`}
              >
                {loadingPlayers ? 'Lade...' : 'Einzelne Liste anzeigen'}
              </button>
              
              <div className="flex gap-2">
                <div className="w-20">
                  <input
                    type="number"
                    min="1"
                    max="50"
                    className="block w-full border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 sm:text-sm border p-2 h-10 text-center"
                    value={topX}
                    onChange={(e) => setTopX(parseInt(e.target.value) || 1)}
                    title="Anzahl der Top-Spieler (X)"
                  />
                </div>
                <button
                  onClick={handleFetchTop5}
                  disabled={loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId}
                  className={`flex-1 flex justify-center py-2 px-4 border border-blue-600 rounded-md shadow-sm text-sm font-medium ${
                    loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId ? 'text-gray-400 border-gray-300' : 'text-blue-600 hover:bg-blue-50'
                  } focus:outline-none h-10`}
                >
                  {loadingTop5 ? `Lade Top-${topX}...` : `Top-${topX} Übersicht`}
                </button>
              </div>
            </div>
          </div>

          <div className="mt-8 pt-6 border-t border-gray-100">
            <h2 className="text-lg font-semibold text-gray-800 mb-4">Turnier-Rückerstattung prüfen</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-end">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Mind. Trainings (K)</label>
                <input
                  type="number"
                  min="0"
                  className="block w-full border-gray-300 rounded-md shadow-sm focus:ring-green-500 focus:border-green-500 sm:text-sm border p-2"
                  value={minTrainings}
                  onChange={(e) => setMinTrainings(parseInt(e.target.value) || 0)}
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Anwesenheitsliste (.xlsx)</label>
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  ref={fileInputRef}
                  onChange={handleFileUpload}
                  className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-green-50 file:text-green-700 hover:file:bg-green-100 cursor-pointer"
                />
              </div>
              <div className="md:col-span-2">
                <button
                  onClick={checkEligibility}
                  disabled={loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId || attendanceData.length === 0}
                  className={`w-full flex justify-center py-3 px-4 border border-transparent rounded-md shadow-sm text-base font-bold text-white ${
                    loadingPlayers || loadingTop5 || checkingEligibility || !selectedClubId || attendanceData.length === 0 ? 'bg-gray-400' : 'bg-green-600 hover:bg-green-700'
                  } focus:outline-none h-12`}
                >
                  {checkingEligibility ? 'Prüfe Berechtigung...' : 'Rückerstattung prüfen (Top-X & K)'}
                </button>
              </div>
            </div>
          </div>

          {error && (
            <div className="mt-4 p-4 bg-red-50 border border-red-200 text-red-700 rounded-md text-sm">
              {error}
            </div>
          )}
        </div>

        {eligibilityResults.length > 0 && (
          <div className="bg-white shadow-sm rounded-lg overflow-hidden border border-green-200">
            <div className="px-4 py-5 sm:px-6 bg-green-50 border-b border-green-200">
              <h3 className="text-lg leading-6 font-bold text-green-800">
                Berechtigte Spieler für Rückerstattung
              </h3>
              <p className="text-sm text-green-600">Top-{topX} im Club und mind. {minTrainings} Trainings</p>
            </div>
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Kategorie</th>
                    <th className="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">Club-Rang</th>
                    <th className="px-6 py-3 text-center text-xs font-medium text-gray-500 uppercase tracking-wider">ÖBV-Rang</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Trainings</th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {eligibilityResults.map((res, idx) => (
                    <tr key={idx} className="hover:bg-green-50">
                      <td className="px-6 py-4 whitespace-nowrap text-sm font-bold text-gray-900">{res.name}</td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-600">{res.category}</td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-center text-blue-600 font-bold">{res.clubRank}.</td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-center text-gray-500">#{res.rank}</td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-right font-mono text-green-600 font-bold">{res.attendance}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {players.length > 0 && (
          <div className="bg-white shadow-sm rounded-lg overflow-hidden border border-gray-200 mt-8">
            <div className="px-4 py-5 sm:px-6 bg-gray-50 border-b border-gray-200">
              <h3 className="text-lg leading-6 font-medium text-gray-900">
                Spieler von {selectedClubName}
              </h3>
            </div>
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-16">Rang</th>
                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-16">Club-Rang</th>
                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Spieler</th>
                    <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Verein</th>
                    <th className="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Punkte</th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {players.map((player, idx) => (
                    <tr key={idx} className="hover:bg-gray-50">
                      <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">{player.rank}</td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-blue-600 font-bold">{idx + 1}.</td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-600">{player.name}</td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{player.club}</td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900 text-right font-mono">{player.points}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {top5Results.length > 0 && (
          <div className="mt-8 space-y-6">
            <h2 className="text-xl font-bold text-gray-900">Top-{topX} Übersicht: {selectedClubName}</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {top5Results.map((result, idx) => (
                <div key={idx} className="bg-white shadow-sm rounded-lg border border-gray-200 overflow-hidden">
                  <div className="px-4 py-3 bg-gray-50 border-b border-gray-200">
                    <h3 className="text-sm font-bold text-gray-700">{result.categoryName}</h3>
                  </div>
                  <div className="p-0">
                    {result.players.length > 0 ? (
                      <ul className="divide-y divide-gray-100">
                        {result.players.map((p, pIdx) => (
                          <li key={pIdx} className="px-4 py-2 flex justify-between text-sm hover:bg-gray-50">
                            <span className="text-gray-900 font-medium">
                              <span className="text-blue-600 font-bold mr-2">{pIdx + 1}.</span>
                              <span className="text-gray-400 mr-2">#{p.rank}</span>
                              {p.name}
                            </span>
                            <span className="text-gray-500 font-mono">{p.points}</span>
                          </li>
                        ))}
                      </ul>
                    ) : (
                      <p className="p-4 text-sm text-gray-400 italic">Keine Spieler gelistet</p>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {!loadingPlayers && !loadingTop5 && !checkingEligibility && selectedClubId && players.length === 0 && top5Results.length === 0 && eligibilityResults.length === 0 && !error && (
          <div className="text-center py-12 bg-white rounded-lg border border-gray-200 shadow-sm">
            <p className="text-gray-500">Wählen Sie eine Aktion oben aus.</p>
          </div>
        )}
      </div>
    </main>
  );
}
