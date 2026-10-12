import { NextResponse } from 'next/server';
import type { NextRequest } from 'next/server';

/**
 * The Electron app starts its Next server on 127.0.0.1. The JSON APIs expose
 * local ranking data and can trigger scans, so they must not become an
 * unauthenticated public web API if someone deploys the server externally.
 * Also block cross-origin browser requests to the loopback service (CSRF/DNS
 * rebinding defense). Local direct clients remain part of the desktop trust
 * boundary; this is not a substitute for user authentication on a hosted app.
 */
export function proxy(request: NextRequest) {
    const hostHeader = request.headers.get('host') || '';
    let hostname = '';

    try {
        hostname = new URL(`http://${hostHeader}`).hostname.toLowerCase();
    } catch {
        return NextResponse.json({ error: 'Forbidden' }, { status: 403 });
    }

    if (hostname !== '127.0.0.1' && hostname !== 'localhost') {
        return NextResponse.json({ error: 'This API is available only through the local desktop app' }, { status: 403 });
    }

    const origin = request.headers.get('origin');
    if (origin) {
        try {
            if (new URL(origin).origin !== request.nextUrl.origin) {
                return NextResponse.json({ error: 'Cross-origin API request denied' }, { status: 403 });
            }
        } catch {
            return NextResponse.json({ error: 'Invalid request origin' }, { status: 403 });
        }
    } else if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method)) {
        return NextResponse.json({ error: 'Origin header required for state-changing requests' }, { status: 403 });
    }

    const fetchSite = request.headers.get('sec-fetch-site');
    if (fetchSite && fetchSite !== 'same-origin' && fetchSite !== 'none') {
        return NextResponse.json({ error: 'Cross-site API request denied' }, { status: 403 });
    }

    return NextResponse.next();
}

export const config = {
    matcher: ['/api/:path*'],
};
