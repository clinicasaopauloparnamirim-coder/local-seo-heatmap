import { prisma } from '@/lib/prisma';
import { NextResponse } from 'next/server';
import { enqueueScan } from '@/lib/scanQueue';
import { logger } from '@/lib/logger';

export async function GET() {
    try {
        const scans = await prisma.scan.findMany({
            orderBy: { createdAt: 'desc' },
        });
        return NextResponse.json({ scans });
    } catch (error) {
        logger.error('Scans GET error', 'SCANNER', { error: String(error) });
        return NextResponse.json({ scans: [], error: 'Failed to fetch scans' }, { status: 500 });
    }
}

export async function POST(req: Request) {
    try {
        const {
            keyword,
            radius,
            gridSize,
            frequency,
            businessName,
            shape,
            customPoints,
            lat,
            lng,
            placeId
        } = await req.json();

        if (!keyword || typeof keyword !== 'string' || !keyword.trim()) {
            return NextResponse.json({ error: 'Keyword is required' }, { status: 400 });
        }
        if (keyword.trim().length > 200) {
            return NextResponse.json({ error: 'Keyword must be 200 characters or less' }, { status: 400 });
        }

        // Never create a scan from a guessed/default location. Coordinates must come
        // from the selected place, geocoding, or deliberate map placement.
        if (
            typeof lat !== 'number' || !Number.isFinite(lat) || lat < -90 || lat > 90 ||
            typeof lng !== 'number' || !Number.isFinite(lng) || lng < -180 || lng > 180
        ) {
            return NextResponse.json({
                error: 'A valid map location is required. Select a place or geocode the address before starting the scan.',
            }, { status: 400 });
        }
        const centerLat = lat;
        const centerLng = lng;

        if (businessName !== undefined && (typeof businessName !== 'string' || businessName.length > 200)) {
            return NextResponse.json({ error: 'Business name must be a string of 200 characters or less' }, { status: 400 });
        }
        if (placeId !== undefined && (typeof placeId !== 'string' || placeId.length > 300)) {
            return NextResponse.json({ error: 'Place ID is invalid' }, { status: 400 });
        }
        if (customPoints !== undefined && customPoints !== null) {
            if (!Array.isArray(customPoints) || customPoints.length > 225) {
                return NextResponse.json({ error: 'Custom grid must contain between 1 and 225 valid points' }, { status: 400 });
            }
            const allPointsValid = customPoints.every((point: unknown) => {
                if (!point || typeof point !== 'object') return false;
                const candidate = point as { lat?: unknown; lng?: unknown };
                return typeof candidate.lat === 'number' && Number.isFinite(candidate.lat) && candidate.lat >= -90 && candidate.lat <= 90 &&
                    typeof candidate.lng === 'number' && Number.isFinite(candidate.lng) && candidate.lng >= -180 && candidate.lng <= 180;
            });
            if (!allPointsValid) return NextResponse.json({ error: 'Custom grid contains invalid coordinates' }, { status: 400 });
        }

        // Validate and clamp bounds
        const parsedRadius = Math.min(Math.max(parseFloat(radius) || 5, 0.5), 100);
        const parsedGridSize = Math.min(Math.max(parseInt(gridSize) || 3, 1), 15);
        const validShapes = ['SQUARE', 'CIRCLE', 'ZIP', 'SMART'];
        const validFrequencies = ['ONCE', 'DAILY', 'WEEKLY'];

        const scan = await prisma.scan.create({
            data: {
                keyword: keyword.trim(),
                centerLat,
                centerLng,
                radius: parsedRadius,
                gridSize: parsedGridSize,
                shape: validShapes.includes(shape) ? shape : 'SQUARE',
                customPoints: customPoints ? JSON.stringify(customPoints) : null,
                frequency: validFrequencies.includes(frequency) ? frequency : 'ONCE',
                businessName: businessName || undefined,
                placeId: placeId || undefined,
                status: 'PENDING',
            },
        });

        // Start scan via queue (respects concurrency limit)
        logger.info(`New scan created: "${keyword}"`, 'API', { scanId: scan.id });
        const queueResult = enqueueScan(scan.id);
        logger.info(`Scan ${scan.id} enqueue result: ${queueResult}`, 'API');

        return NextResponse.json(scan);
    } catch (error) {
        logger.error('Scan creation failed', 'SCANNER', { error: String(error), stack: error instanceof Error ? error.stack : undefined });
        return NextResponse.json({ error: 'Failed to create scan' }, { status: 500 });
    }
}
