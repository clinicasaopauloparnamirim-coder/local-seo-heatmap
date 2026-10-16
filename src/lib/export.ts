import ExcelJS from 'exceljs';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

interface ScanData {
    id: string;
    keyword: string;
    status: string;
    gridSize: number;
    radius: number;
    frequency: string;
    createdAt: string;
    centerLat: number;
    centerLng: number;
    businessName?: string;
    results: any[];
}

interface CompetitorSummary {
    name: string;
    appearances: number;
    avgRank: number;
    bestRank: number;
    rating?: number;
    reviews?: number;
    address?: string;
}

// Helper to aggregate competitor data
function aggregateCompetitors(results: any[]): CompetitorSummary[] {
    const competitorMap = new Map<string, { ranks: number[]; rating?: number; reviews?: number; address?: string }>();

    results.forEach(point => {
        try {
            const businesses = JSON.parse(point.topResults);
            businesses.forEach((biz: any) => {
                const name = biz.name?.toLowerCase().trim();
                if (!name) return;

                const existing = competitorMap.get(name) || { ranks: [], rating: undefined, reviews: undefined, address: undefined };
                existing.ranks.push(biz.rank || 21);
                if (biz.rating) existing.rating = biz.rating;
                if (biz.reviews) existing.reviews = biz.reviews;
                if (biz.address) existing.address = biz.address;
                competitorMap.set(name, existing);
            });
        } catch (e) { }
    });

    const competitors: CompetitorSummary[] = [];
    competitorMap.forEach((data, name) => {
        competitors.push({
            name: name.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' '),
            appearances: data.ranks.length,
            avgRank: data.ranks.reduce((a, b) => a + b, 0) / data.ranks.length,
            bestRank: Math.min(...data.ranks),
            rating: data.rating,
            reviews: data.reviews,
            address: data.address
        });
    });

    return competitors.sort((a, b) => a.avgRank - b.avgRank);
}

// Helper to calculate metrics
function calculateMetrics(results: any[], businessName?: string) {
    const totalPoints = results.length;
    let rankedPoints = 0;
    let top3Points = 0;
    let top10Points = 0;
    let totalRank = 0;

    results.forEach(r => {
        if (r.rank) {
            rankedPoints++;
            totalRank += r.rank;
            if (r.rank <= 3) top3Points++;
            if (r.rank <= 10) top10Points++;
        }
    });

    const avgRank = rankedPoints > 0 ? totalRank / rankedPoints : 0;
    const top3Percentage = totalPoints > 0 ? (top3Points / totalPoints) * 100 : 0;
    const top10Percentage = totalPoints > 0 ? (top10Points / totalPoints) * 100 : 0;
    const visibilityScore = totalPoints > 0 ? (rankedPoints / totalPoints) * 100 : 0;

    return {
        totalPoints,
        rankedPoints,
        top3Points,
        top10Points,
        avgRank,
        top3Percentage,
        top10Percentage,
        visibilityScore
    };
}

export async function exportToXLSX(scanName: string, data: any[]) {
    const workbook = new ExcelJS.Workbook();
    addScanToWorkbook(workbook, scanName, data);

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${scanName.replace(/\s+/g, '_')}_results.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
}

function addScanToWorkbook(workbook: ExcelJS.Workbook, scanName: string, data: any[]) {
    const sheetName = scanName.substring(0, 31).replace(/[\\\/\?\*\[\]]/g, '_');
    const worksheet = workbook.addWorksheet(sheetName);

    worksheet.columns = [
        { header: 'Rank', key: 'rank', width: 10 },
        { header: 'Business Name', key: 'name', width: 40 },
        { header: 'Rating', key: 'rating', width: 10 },
        { header: 'Reviews', key: 'reviews', width: 10 },
        { header: 'Address', key: 'address', width: 60 },
        { header: 'URL', key: 'url', width: 80 },
        { header: 'Point Lat', key: 'lat', width: 15 },
        { header: 'Point Lng', key: 'lng', width: 15 },
    ];

    data.forEach(point => {
        try {
            const results = JSON.parse(point.topResults);
            results.forEach((res: any) => {
                worksheet.addRow({
                    rank: res.rank,
                    name: res.name,
                    rating: res.rating || 'N/A',
                    reviews: res.reviews || 0,
                    address: res.address || '',
                    url: res.url || '',
                    lat: point.lat,
                    lng: point.lng
                });
            });
        } catch (e) {
            console.error('Failed to parse results for excel export', e);
        }
    });
}

export async function exportAllScansToXLSX(scans: any[]) {
    const workbook = new ExcelJS.Workbook();

    // Summary sheet first
    const summarySheet = workbook.addWorksheet('All Scans Summary');
    summarySheet.columns = [
        { header: 'Keyword', key: 'keyword', width: 30 },
        { header: 'Status', key: 'status', width: 15 },
        { header: 'Grid Size', key: 'gridSize', width: 10 },
        { header: 'Radius', key: 'radius', width: 10 },
        { header: 'Created At', key: 'createdAt', width: 25 },
        { header: 'Center Lat', key: 'centerLat', width: 15 },
        { header: 'Center Lng', key: 'centerLng', width: 15 },
    ];

    scans.forEach(scan => {
        summarySheet.addRow({
            keyword: scan.keyword,
            status: scan.status,
            gridSize: `${scan.gridSize}x${scan.gridSize}`,
            radius: `${scan.radius}km`,
            createdAt: new Date(scan.createdAt).toLocaleString(),
            centerLat: scan.centerLat,
            centerLng: scan.centerLng
        });

        if (scan.results && scan.results.length > 0) {
            addScanToWorkbook(workbook, scan.keyword, scan.results);
        }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `GBP Rank Tracker_All_Scans_${new Date().toISOString().split('T')[0]}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
}

function parseTopResultsForExport(topResults: unknown): any[] {
    if (Array.isArray(topResults)) return topResults;
    if (typeof topResults !== 'string' || !topResults.trim()) return [];
    try {
        const parsed = JSON.parse(topResults);
        return Array.isArray(parsed) ? parsed : [];
    } catch {
        return [];
    }
}

function buildExportRows(scanName: string, data: any[], scan?: Partial<ScanData>) {
    const rows: Record<string, string | number | boolean | null>[] = [];
    data.forEach((point, pointIndex) => {
        const businesses = parseTopResultsForExport(point?.topResults);
        const common = {
            scan_keyword: scan?.keyword || scanName,
            business_tracked: scan?.businessName || '',
            scan_id: scan?.id || '',
            scan_status: scan?.status || '',
            scan_created_at: scan?.createdAt || '',
            grid_size: scan?.gridSize ?? '',
            radius_km: scan?.radius ?? '',
            center_latitude: scan?.centerLat ?? '',
            center_longitude: scan?.centerLng ?? '',
            grid_point_index: pointIndex + 1,
            point_latitude: point?.lat ?? '',
            point_longitude: point?.lng ?? '',
            target_rank: point?.rank ?? '',
            run_id: point?.runId || '',
            run_at: point?.runAt || point?.capturedAt || '',
        };
        if (businesses.length === 0) {
            rows.push({
                ...common,
                result_rank: '',
                result_name: '',
                place_id: '',
                cid: '',
                rating: '',
                review_count: '',
                category: '',
                address: '',
                phone: '',
                website: '',
            });
            return;
        }
        businesses.forEach((business: any) => {
            rows.push({
                ...common,
                result_rank: business?.rank ?? '',
                result_name: business?.name ?? '',
                place_id: business?.placeId ?? business?.place_id ?? '',
                cid: business?.cid ?? '',
                rating: business?.rating ?? '',
                review_count: business?.reviews ?? business?.reviewCount ?? '',
                category: business?.category ?? business?.type ?? '',
                address: business?.address ?? '',
                phone: business?.phone ?? '',
                website: business?.website ?? business?.url ?? '',
            });
        });
    });
    return rows;
}

function csvCell(value: unknown): string {
    const raw = value === null || value === undefined ? '' : String(value);
    const safe = /^[=+@]/.test(raw) || /^-\\s/.test(raw) ? "'" + raw : raw;
    return '"' + safe.replace(/"/g, '""') + '"';
}

function triggerExportDownload(filename: string, blob: Blob): void {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => window.URL.revokeObjectURL(url), 1000);
}

export function exportToCSV(scanName: string, data: any[], scan?: Partial<ScanData>): void {
    const rows = buildExportRows(scanName, data, scan);
    const headers = rows.length ? Object.keys(rows[0]) : [
        'scan_keyword', 'business_tracked', 'scan_id', 'scan_status', 'scan_created_at',
        'grid_size', 'radius_km', 'center_latitude', 'center_longitude', 'grid_point_index',
        'point_latitude', 'point_longitude', 'target_rank', 'run_id', 'run_at', 'result_rank',
        'result_name', 'place_id', 'cid', 'rating', 'review_count', 'category', 'address',
        'phone', 'website',
    ];
    const csv = [headers.map(csvCell).join(','), ...rows.map(row => headers.map(header => csvCell(row[header])).join(','))].join('\\r\\n');
    const filename = (scanName || 'heatmap').trim().replace(/[^\\p{L}\\p{N}_-]+/gu, '_') || 'heatmap';
    triggerExportDownload(filename + '_results.csv', new Blob(['\\uFEFF', csv], { type: 'text/csv;charset=utf-8;' }));
}

export function exportToJSON(scanName: string, data: any[], scan?: Partial<ScanData>): void {
    const points = data.map((point, index) => ({
        pointIndex: index + 1,
        id: point?.id ?? null,
        latitude: point?.lat ?? null,
        longitude: point?.lng ?? null,
        targetRank: point?.rank ?? null,
        targetName: point?.targetName ?? scan?.businessName ?? null,
        placeId: point?.placeId ?? null,
        cid: point?.cid ?? null,
        runId: point?.runId ?? null,
        runAt: point?.runAt ?? point?.capturedAt ?? null,
        businesses: parseTopResultsForExport(point?.topResults),
    }));
    const payload = {
        format: 'heartseomap-scan-export',
        formatVersion: 1,
        exportedAt: new Date().toISOString(),
        scan: {
            id: scan?.id ?? null,
            keyword: scan?.keyword ?? scanName,
            status: scan?.status ?? null,
            businessName: scan?.businessName ?? null,
            gridSize: scan?.gridSize ?? null,
            radiusKm: scan?.radius ?? null,
            frequency: scan?.frequency ?? null,
            createdAt: scan?.createdAt ?? null,
            center: {
                latitude: scan?.centerLat ?? null,
                longitude: scan?.centerLng ?? null,
            },
        },
        totalGridPoints: points.length,
        totalBusinessRows: points.reduce((sum, point) => sum + point.businesses.length, 0),
        points,
    };
    const filename = (scanName || 'heatmap').trim().replace(/[^\\p{L}\\p{N}_-]+/gu, '_') || 'heatmap';
    triggerExportDownload(filename + '_results.json', new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8;' }));
}

/**
 * Enhanced PDF Export - Comprehensive Visual Report
 */
export async function exportToPDF(scanName: string, data: any[], scan?: ScanData) {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 14;
    let yPos = 0;

    // Colors
    const primaryColor: [number, number, number] = [37, 99, 235]; // Blue
    const successColor: [number, number, number] = [34, 197, 94]; // Green
    const warningColor: [number, number, number] = [245, 158, 11]; // Amber
    const dangerColor: [number, number, number] = [239, 68, 68]; // Red
    const grayColor: [number, number, number] = [107, 114, 128];
    const darkColor: [number, number, number] = [17, 24, 39];

    // ==========================================
    // PAGE 1: TITLE PAGE
    // ==========================================

    // Header background
    doc.setFillColor(...primaryColor);
    doc.rect(0, 0, pageWidth, 70, 'F');

    // Logo/Brand
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(32);
    doc.setFont('helvetica', 'bold');
    doc.text('GBP RANK TRACKER', margin, 35);

    doc.setFontSize(12);
    doc.setFont('helvetica', 'normal');
    doc.text('Local SEO Intelligence Report', margin, 50);

    // Report Title
    yPos = 90;
    doc.setTextColor(...darkColor);
    doc.setFontSize(24);
    doc.setFont('helvetica', 'bold');
    doc.text(`Keyword: "${scanName}"`, margin, yPos);

    // Scan Details Box
    yPos = 110;
    doc.setFillColor(248, 250, 252);
    doc.roundedRect(margin, yPos, pageWidth - margin * 2, 50, 3, 3, 'F');

    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(...grayColor);
    doc.text('SCAN DETAILS', margin + 10, yPos + 12);

    doc.setFont('helvetica', 'normal');
    doc.setTextColor(...darkColor);
    const detailsY = yPos + 25;

    // Left column
    doc.text(`Generated: ${new Date().toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}`, margin + 10, detailsY);
    doc.text(`Grid Points: ${data.length} locations scanned`, margin + 10, detailsY + 10);

    // Right column
    if (scan) {
        doc.text(`Grid Size: ${scan.gridSize}×${scan.gridSize}`, pageWidth / 2 + 10, detailsY);
        doc.text(`Radius: ${scan.radius} km`, pageWidth / 2 + 10, detailsY + 10);
        doc.text(`Center: ${scan.centerLat.toFixed(4)}, ${scan.centerLng.toFixed(4)}`, pageWidth / 2 + 10, detailsY + 20);
    }

    // Target business if specified
    if (scan?.businessName) {
        yPos = 175;
        doc.setFillColor(...primaryColor);
        doc.roundedRect(margin, yPos, pageWidth - margin * 2, 25, 3, 3, 'F');
        doc.setTextColor(255, 255, 255);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(11);
        doc.text(`🎯 Tracking: ${scan.businessName}`, margin + 10, yPos + 16);
    }

    // ==========================================
    // PAGE 1: EXECUTIVE SUMMARY
    // ==========================================

    const metrics = calculateMetrics(data, scan?.businessName);

    yPos = scan?.businessName ? 215 : 180;
    doc.setTextColor(...darkColor);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('Executive Summary', margin, yPos);

    // Metric boxes
    yPos += 15;
    const boxWidth = (pageWidth - margin * 2 - 30) / 4;
    const boxHeight = 35;

    const metricBoxes = [
        { label: 'Grid Points', value: metrics.totalPoints.toString(), color: primaryColor },
        { label: 'Avg Rank', value: metrics.avgRank > 0 ? `#${metrics.avgRank.toFixed(1)}` : 'N/A', color: metrics.avgRank <= 5 ? successColor : metrics.avgRank <= 10 ? warningColor : dangerColor },
        { label: 'Top 3 Rate', value: `${metrics.top3Percentage.toFixed(0)}%`, color: metrics.top3Percentage >= 50 ? successColor : metrics.top3Percentage >= 25 ? warningColor : dangerColor },
        { label: 'Visibility', value: `${metrics.visibilityScore.toFixed(0)}%`, color: metrics.visibilityScore >= 70 ? successColor : metrics.visibilityScore >= 40 ? warningColor : dangerColor }
    ];

    metricBoxes.forEach((box, i) => {
        const x = margin + i * (boxWidth + 10);
        doc.setFillColor(...box.color);
        doc.roundedRect(x, yPos, boxWidth, boxHeight, 3, 3, 'F');

        doc.setTextColor(255, 255, 255);
        doc.setFontSize(18);
        doc.setFont('helvetica', 'bold');
        doc.text(box.value, x + boxWidth / 2, yPos + 18, { align: 'center' });

        doc.setFontSize(8);
        doc.setFont('helvetica', 'normal');
        doc.text(box.label.toUpperCase(), x + boxWidth / 2, yPos + 28, { align: 'center' });
    });

    // ==========================================
    // PAGE 2: TOP COMPETITORS
    // ==========================================

    doc.addPage();
    yPos = 20;

    // Header
    doc.setFillColor(...primaryColor);
    doc.rect(0, 0, pageWidth, 40, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(18);
    doc.setFont('helvetica', 'bold');
    doc.text('Top Competitors Analysis', margin, 25);

    yPos = 55;
    doc.setTextColor(...darkColor);
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text('Market Leaders by Average Rank', margin, yPos);

    // Get top competitors
    const competitors = aggregateCompetitors(data);
    const topCompetitors = competitors.slice(0, 10);

    // Competitor cards
    yPos += 10;
    topCompetitors.forEach((comp, idx) => {
        const cardY = yPos + idx * 18;

        // Rank badge
        const rankColor = idx < 3 ? successColor : idx < 7 ? warningColor : dangerColor;
        doc.setFillColor(...rankColor);
        doc.roundedRect(margin, cardY, 20, 14, 2, 2, 'F');
        doc.setTextColor(255, 255, 255);
        doc.setFontSize(10);
        doc.setFont('helvetica', 'bold');
        doc.text(`#${idx + 1}`, margin + 10, cardY + 9, { align: 'center' });

        // Business name
        doc.setTextColor(...darkColor);
        doc.setFontSize(10);
        doc.setFont('helvetica', 'bold');
        const displayName = comp.name.length > 35 ? comp.name.substring(0, 35) + '...' : comp.name;
        doc.text(displayName, margin + 25, cardY + 9);

        // Stats
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(...grayColor);
        doc.setFontSize(8);

        const stats = [];
        stats.push(`Avg: #${comp.avgRank.toFixed(1)}`);
        stats.push(`Best: #${comp.bestRank}`);
        stats.push(`Seen: ${comp.appearances}x`);
        if (comp.rating) stats.push(`★${comp.rating}`);
        if (comp.reviews) stats.push(`${comp.reviews} reviews`);

        doc.text(stats.join('  |  '), pageWidth - margin, cardY + 9, { align: 'right' });
    });

    // Market Share Insights
    yPos = yPos + topCompetitors.length * 18 + 20;
    doc.setFillColor(248, 250, 252);
    doc.roundedRect(margin, yPos, pageWidth - margin * 2, 45, 3, 3, 'F');

    doc.setTextColor(...primaryColor);
    doc.setFontSize(11);
    doc.setFont('helvetica', 'bold');
    doc.text('💡 Market Insights', margin + 10, yPos + 15);

    doc.setTextColor(...darkColor);
    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');

    const insights = [];
    if (topCompetitors[0]) {
        insights.push(`• The market leader "${topCompetitors[0].name}" appears in ${topCompetitors[0].appearances} of ${data.length} grid points.`);
    }
    if (metrics.top3Percentage < 30) {
        insights.push(`• Low Top-3 visibility (${metrics.top3Percentage.toFixed(0)}%) indicates strong competition in this market.`);
    } else if (metrics.top3Percentage > 60) {
        insights.push(`• High Top-3 visibility (${metrics.top3Percentage.toFixed(0)}%) shows market dominance opportunities.`);
    }
    insights.push(`• ${competitors.length} unique businesses compete for visibility in this area.`);

    insights.slice(0, 3).forEach((insight, i) => {
        doc.text(insight, margin + 10, yPos + 25 + i * 8);
    });

    // ==========================================
    // PAGE 3: DETAILED RANKINGS TABLE
    // ==========================================

    doc.addPage();

    // Header
    doc.setFillColor(...primaryColor);
    doc.rect(0, 0, pageWidth, 40, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(18);
    doc.setFont('helvetica', 'bold');
    doc.text('Grid Point Breakdown', margin, 25);

    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.text(`Detailed rankings from ${data.length} scan points`, margin, 35);

    // Build table data - show top 3 from each point
    const tableData: any[][] = [];
    data.forEach((point, pointIdx) => {
        try {
            const results = JSON.parse(point.topResults);
            const top3 = results.slice(0, 3);
            top3.forEach((res: any, resIdx: number) => {
                tableData.push([
                    resIdx === 0 ? `Point ${pointIdx + 1}` : '',
                    resIdx === 0 ? `${point.lat.toFixed(4)}, ${point.lng.toFixed(4)}` : '',
                    `#${res.rank}`,
                    res.name?.substring(0, 30) || 'Unknown',
                    res.rating ? `★${res.rating}` : '-',
                    res.reviews || '-'
                ]);
            });
        } catch (e) { }
    });

    autoTable(doc, {
        head: [['Point', 'Coordinates', 'Rank', 'Business Name', 'Rating', 'Reviews']],
        body: tableData.slice(0, 60), // Limit to prevent overflow
        startY: 50,
        theme: 'grid',
        headStyles: { fillColor: primaryColor, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
        styles: { fontSize: 7, cellPadding: 2 },
        columnStyles: {
            0: { cellWidth: 20, fontStyle: 'bold' },
            1: { cellWidth: 35 },
            2: { cellWidth: 15, halign: 'center' },
            3: { cellWidth: 60 },
            4: { cellWidth: 20, halign: 'center' },
            5: { cellWidth: 20, halign: 'center' }
        },
        alternateRowStyles: { fillColor: [248, 250, 252] }
    });

    // ==========================================
    // FOOTER on last page
    // ==========================================

    const lastPageY = (doc as any).lastAutoTable?.finalY || 200;
    if (lastPageY < pageHeight - 40) {
        doc.setFillColor(248, 250, 252);
        doc.rect(0, pageHeight - 25, pageWidth, 25, 'F');

        doc.setTextColor(...grayColor);
        doc.setFontSize(8);
        doc.setFont('helvetica', 'normal');
        doc.text('Generated by GBP Rank Tracker - Local SEO Intelligence Grid', margin, pageHeight - 12);
        doc.text(`Report ID: ${Date.now().toString(36).toUpperCase()}`, pageWidth - margin, pageHeight - 12, { align: 'right' });
    }

    // Save
    doc.save(`${scanName.replace(/\s+/g, '_')}_GBP Rank Tracker_Report.pdf`);
}

