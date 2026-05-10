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
  rank: string;
  name: string;
  club: string;
  points: string;
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

  // 1. POST to cookiewall/Save
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

  // Extract cookies from response
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

  // Find actual ranking ID
  $('a[href*="category.aspx"]').each((i, el) => {
    const href = $(el).attr('href');
    if (href && href.includes('id=')) {
      const match = href.match(/id=([^&]+)/);
      if (match) {
        actualId = match[1];
        return false; // break
      }
    }
  });

  // Go to category page to get pubs and clubs
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
  let url = `${BASE_URL}/ranking/category.aspx?id=${actualId}&category=${catId}&ps=500`;
  if (pubId) {
    url += `&publicationid=${pubId}`;
  }
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
      const pName = $(cols[3]).text().trim();
      const cName = $(cols[10]).text().trim();
      
      if (pName) {
        players.push({
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
