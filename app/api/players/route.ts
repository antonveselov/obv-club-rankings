import { NextRequest, NextResponse } from 'next/server';
import { fetchPlayers } from '@/lib/scraper';

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const actualId = searchParams.get('actualId');
  const catId = searchParams.get('catId');
  const pubId = searchParams.get('pubId') || undefined;
  const clubId = searchParams.get('clubId') || undefined;

  if (!actualId || !catId) {
    return NextResponse.json({ error: 'Missing actualId or catId' }, { status: 400 });
  }

  try {
    const players = await fetchPlayers(actualId, catId, pubId, clubId);
    return NextResponse.json(players);
  } catch (error: any) {
    console.error('Players fetch error:', error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
