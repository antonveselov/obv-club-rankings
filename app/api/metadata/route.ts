import { NextResponse } from 'next/server';
import { fetchMetadata } from '@/lib/scraper';

// Simple in-memory cache for metadata (expires in 1 hour)
let cachedMetadata: any = null;
let lastFetchTime = 0;
const CACHE_TTL = 3600000; // 1 hour

export async function GET() {
  const now = Date.now();
  if (cachedMetadata && (now - lastFetchTime < CACHE_TTL)) {
    return NextResponse.json(cachedMetadata);
  }

  try {
    const metadata = await fetchMetadata();
    cachedMetadata = metadata;
    lastFetchTime = now;
    return NextResponse.json(metadata);
  } catch (error: any) {
    console.error('Metadata fetch error:', error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
