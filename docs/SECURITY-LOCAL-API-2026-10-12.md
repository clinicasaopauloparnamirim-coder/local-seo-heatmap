# Local SEO Heatmap — security fixes (2026-10-12)

## Request boundary
The desktop runtime binds Next.js to `127.0.0.1`. The new Next.js 16 `src/proxy.ts` adds a second boundary for every `/api/*` handler: only `localhost`/loopback Host values, same-origin requests, and same-origin Fetch Metadata are accepted. State-changing browser requests must include a matching Origin. Cross-origin and rebinding-style requests receive HTTP 403.

This protects the local desktop API from browser-based cross-site access. It does **not** provide identity/authentication for a deliberately hosted web deployment; hosted deployment remains unsupported and fail-closed.

## Location correctness
New scans no longer silently fall back to Chicago. Latitude/longitude must be finite and within geographic bounds. Custom grids must contain valid coordinate objects and are capped at 225 points. The creation UI defaults to the Clínica São Paulo address in Centro de Parnamirim/RN and invalidates the location confirmation when the address is edited. The user must geocode the new address, select a place with coordinates, use a resolved city grid, or deliberately move the map before proceeding.

## Input/output hygiene
Scan PATCH inputs are validated and limited to supported fields/ranges. API errors no longer disclose raw database/exception details to the renderer.

## Build reliability
The Next.js `typescript.ignoreBuildErrors` escape hatch has been removed. A new pull-request CI verifies Next build/type checking, Electron compilation, and package invariants before changes are accepted.
