import { NextRequest, NextResponse } from 'next/server';
import { fetchPlayerHistory, fetchMetadata } from '@/lib/scraper';

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const actualId = searchParams.get('actualId');
  const playerId = searchParams.get('playerId');
  const clubId = searchParams.get('clubId') || '';

  if (!actualId || !playerId) {
    return NextResponse.json({ error: 'Missing actualId or playerId' }, { status: 400 });
  }

  try {
    // We need the publications list to know which dates to fetch
    // To save time, we could fetch them here or pass them from frontend
    // Let's fetch them here using our cached metadata logic if possible, 
    // but for now just call fetchMetadata again.
    const meta = await fetchMetadata();
    const history = await fetchPlayerHistory(actualId, playerId, clubId, meta.publications);
    return NextResponse.json(history);
  } catch (error: any) {
    console.error('History fetch error:', error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
