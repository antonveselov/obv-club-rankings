import axios from 'axios';
import * as cheerio from 'cheerio';

const BASE_URL = 'https://obv.tournamentsoftware.com';
const RANKING_URL = `${BASE_URL}/ranking/ranking.aspx?rid=355`;

export interface Publication {
  id: string;
  name: string;
}

export interface Category {
  id: string;
  name: string;
}

export interface Club {
  id: string;
  name: string;
}

export interface Player {
  id: string;
  rank: string;
  name: string;
  club: string;
  points: string;
}

export interface RankHistoryPoint {
  date: string;
  rank: number | null;
}

export interface PlayerHistory {
  singles: RankHistoryPoint[];
  doubles: RankHistoryPoint[];
  mixed: RankHistoryPoint[];
}

export interface SearchResult {
  id: string;
  name: string;
  club: string;
  rankings: {
    singles?: number;
    doubles?: number;
    mixed?: number;
  };
}

export interface Metadata {
  actualId: string;
  publications: Publication[];
  categories: Category[];
  clubs: Club[];
}

export async function getSessionCookies() {
  const session = axios.create({
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    },
    withCredentials: true,
  });

  const params = new URLSearchParams();
  params.append('CookiePurposes', '1');
  params.append('CookiePurposes', '2');
  params.append('CookiePurposes', '4');
  params.append('CookiePurposes', '16');
  params.append('ReturnUrl', '/ranking/ranking.aspx?rid=355');

  const res = await session.post(`${BASE_URL}/cookiewall/Save`, 
    params.toString(),
    {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      maxRedirects: 0,
      validateStatus: (status) => status >= 200 && status < 400,
    }
  );

  const cookies = res.headers['set-cookie'] || [];
  return cookies.map(c => c.split(';')[0]).join('; ');
}

export async function fetchMetadata(): Promise<Metadata> {
  const cookies = await getSessionCookies();
  
  const res = await axios.get(RANKING_URL, {
    headers: {
      'Cookie': cookies,
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    }
  });

  const $ = cheerio.load(res.data);
  let actualId = '355';

  $('a[href*="category.aspx"]').each((i, el) => {
    const href = $(el).attr('href');
    if (href && href.includes('id=')) {
      const match = href.match(/id=([^&]+)/);
      if (match) {
        actualId = match[1];
        return false;
      }
    }
  });

  const catUrl = `${BASE_URL}/ranking/category.aspx?id=${actualId}&category=4670`;
  const resCat = await axios.get(catUrl, {
    headers: {
      'Cookie': cookies,
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    }
  });

  const $cat = cheerio.load(resCat.data);
  
  const publications: Publication[] = [];
  $cat('select.publication option, select[id*="dlPublication"] option').each((i, el) => {
    const id = $cat(el).val() as string;
    const name = $cat(el).text().trim();
    if (id) publications.push({ id, name });
  });

  const clubs: Club[] = [];
  $cat('select[id*="FOG_3"] option').each((i, el) => {
    const id = $cat(el).val() as string;
    const name = $cat(el).text().trim();
    if (id && name && name !== 'All') {
      clubs.push({ id, name });
    }
  });

  const categories: Category[] = [];
  $('a[href*="category.aspx"]').each((i, el) => {
    const href = $(el).attr('href');
    if (href && href.includes(`id=${actualId}`) && href.includes('category=')) {
      const catIdMatch = href.match(/category=([^&]+)/);
      const name = $(el).text().trim();
      if (catIdMatch && name && name !== 'mehr') {
        const catId = catIdMatch[1];
        if (!categories.find(c => c.id === catId)) {
          categories.push({ id: catId, name });
        }
      }
    }
  });

  return {
    actualId,
    publications,
    categories,
    clubs: clubs.sort((a, b) => a.name.localeCompare(b.name))
  };
}

export async function fetchPlayers(
  actualId: string, 
  catId: string, 
  pubId?: string, 
  clubId?: string
): Promise<Player[]> {
  const cookies = await getSessionCookies();
  let url = `${BASE_URL}/ranking/category.aspx?id=${pubId || actualId}&category=${catId}&ps=500`;
  if (clubId) {
    url += `&C${catId}FOG_3_F2048=${clubId}`;
  }

  const res = await axios.get(url, {
    headers: {
      'Cookie': cookies,
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    }
  });

  const $ = cheerio.load(res.data);
  const table = $('table.ruler');
  const players: Player[] = [];

  if (!table.length) return [];

  const rows = table.find('tr');
  let foundHeader = false;
  let headerRowIndex = -1;

  rows.each((i, row) => {
    const text = $(row).text().toLowerCase();
    if (text.includes('spieler') || text.includes('verein')) {
      headerRowIndex = i;
      foundHeader = true;
      return false;
    }
  });

  if (!foundHeader) return [];

  rows.each((i, row) => {
    if (i <= headerRowIndex) return;
    const cols = $(row).find('td');
    if (cols.length >= 11) {
      const pNameLink = $(cols[3]).find('a');
      const pName = $(cols[3]).text().trim();
      const cName = $(cols[10]).text().trim();
      
      let pId = '';
      const href = pNameLink.attr('href');
      if (href) {
        const match = href.match(/player=([^&]+)/);
        if (match) pId = match[1];
      }
      
      if (pName) {
        players.push({
          id: pId,
          rank: $(cols[0]).text().trim(),
          name: pName,
          club: cName,
          points: $(cols[6]).text().trim()
        });
      }
    }
  });

  return players;
}

export async function searchPlayers(query: string): Promise<SearchResult[]> {
  const cookies = await getSessionCookies();
  const url = `${BASE_URL}/find/player?q=${encodeURIComponent(query)}`;
  
  const res = await axios.get(url, {
    headers: {
      'Cookie': cookies,
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    }
  });

  const $ = cheerio.load(res.data);
  const results: SearchResult[] = [];

  // TS search results are often in list__item
  $('.list__item').each((i, el) => {
    const nameLink = $(el).find('a[href*="/player-profile/"]');
    if (nameLink.length) {
      const name = nameLink.text().trim();
      const href = nameLink.attr('href');
      const profileId = href?.split('/player-profile/')[1] || '';
      
      const clubSubtitle = $(el).find('.nav-link__subtitle').text().trim();
      
      if (name && profileId) {
        // Avoid duplicates (TS sometimes shows multiples)
        if (!results.find(r => r.id === profileId)) {
          results.push({
            id: profileId,
            name,
            club: clubSubtitle || 'Unknown Club',
            rankings: {}
          });
        }
      }
    }
  });

  // For each result, let's try to get current rankings
  for (const r of results) {
    try {
      const rankUrl = `${BASE_URL}/player-profile/${r.id}/ranking`;
      const rankRes = await axios.get(rankUrl, {
        headers: {
          'Cookie': cookies,
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        }
      });
      const $rank = cheerio.load(rankRes.data);
      const table = $rank('table.ruler').first();
      
      const catNames = {
        singles: ["singles", "einzel"],
        mixed: ["mixed"],
        doubles: ["doubles", "doppel"]
      };

      table.find('tr').each((i, row) => {
        const cols = $rank(row).find('td');
        if (cols.length >= 2) {
          const catName = $rank(cols[0]).text().toLowerCase();
          const rankText = $rank(cols[1]).text().trim().replace(/[^0-9]/g, '');
          const rank = parseInt(rankText);
          
          if (!isNaN(rank)) {
            if (catNames.mixed.some(n => catName.includes(n))) r.rankings.mixed = rank;
            else if (catNames.singles.some(n => catName.includes(n))) r.rankings.singles = rank;
            else if (catNames.doubles.some(n => catName.includes(n))) r.rankings.doubles = rank;
          }
        }
      });
    } catch (err) {
      console.error(`Error fetching summary for ${r.name}:`, err);
    }
  }

  return results;
}

export async function fetchPlayerHistory(
  actualId: string,
  playerId: string,
  clubId: string,
  publications: Publication[]
): Promise<PlayerHistory> {
  const cookies = await getSessionCookies();
  
  const monthlySubset: Publication[] = [];
  for (let i = 0; i < publications.length && monthlySubset.length < 13; i += 4) {
    monthlySubset.push(publications[i]);
  }
  
  const history: PlayerHistory = {
    singles: [],
    doubles: [],
    mixed: []
  };

  const catNames = {
    singles: ["singles", "einzel"],
    mixed: ["mixed"],
    doubles: ["doubles", "doppel"]
  };

  for (const pub of monthlySubset.reverse()) {
    try {
      const url = `${BASE_URL}/ranking/player.aspx?id=${pub.id}&player=${playerId}`;
      const res = await axios.get(url, {
        headers: {
          'Cookie': cookies,
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        }
      });
      
      const $ = cheerio.load(res.data);
      const table = $('table.ruler').first();
      const rows = table.find('tr');
      
      let sRank: number | null = null;
      let dRank: number | null = null;
      let mRank: number | null = null;

      rows.each((i, row) => {
        const cols = $(row).find('td');
        if (cols.length >= 2) {
          const catName = $(cols[0]).text().toLowerCase();
          const rankText = $(cols[1]).text().trim().replace(/[^0-9]/g, '');
          const rank = parseInt(rankText);
          
          if (!isNaN(rank)) {
            if (catNames.mixed.some(n => catName.includes(n))) mRank = rank;
            else if (catNames.singles.some(n => catName.includes(n))) sRank = rank;
            else if (catNames.doubles.some(n => catName.includes(n))) dRank = rank;
          }
        }
      });

      const dateLabel = pub.name.split(' (')[0]; 
      history.singles.push({ date: dateLabel, rank: sRank });
      history.doubles.push({ date: dateLabel, rank: dRank });
      history.mixed.push({ date: dateLabel, rank: mRank });

    } catch (err) {
      console.error(`Error fetching history for ${pub.name}:`, err);
    }
  }

  return history;
}
