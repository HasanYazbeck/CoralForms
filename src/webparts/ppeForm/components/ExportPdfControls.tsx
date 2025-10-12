import * as React from 'react';
import { DefaultButton } from '@fluentui/react';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';


export type ExportPdfControlsProps = {
    // The root element to capture (your PpeForm containerRef)
    targetRef: React.RefObject<HTMLElement>;
    // For filename
    employeeName?: string;
    fileName?: string;
    // Parent-managed export mode toggle so PpeForm can enable controls while exporting
    exportMode: boolean;
    onExportModeChange: (on: boolean) => void;
    // Optional UX gates/flags
    isClosedBySystem?: boolean;
    disabled?: boolean;
    showButton?: boolean;
    // Errors to host banner/UI
    onError?: (message: string) => void;
    onBusyChange?: (busy: boolean) => void; // NEW
    pdfMarginMm?: number;
};

const ExportPdfControls: React.FC<ExportPdfControlsProps> = ({
    targetRef,
    employeeName,
    fileName,
    exportMode,
    onExportModeChange,
    isClosedBySystem = true,
    disabled,
    showButton = true,
    onError,
    onBusyChange,
    pdfMarginMm = 12,
}) => {

    const exportPdf = React.useCallback(async () => {
        try {
            onBusyChange?.(true);
            onExportModeChange?.(true);
            // Wait a tick for exportMode re-render (ItemsSummaryStack etc.)
            await new Promise(r => setTimeout(r, 50));
            if (!targetRef.current) {
                onError?.('Nothing to export.');
                return;
            }
            if (!isClosedBySystem) {
                onError?.('Only Closed by System forms can be exported.');
                return;
            }

            const pdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' });
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = pdf.internal.pageSize.getHeight();
            const margin = pdfMarginMm;        // mm
            const topMargin = margin;          // you can customize differently if you want
            const bottomMargin = margin;
            const contentWidth = pageWidth - 2 * margin;
            let cursorY = topMargin;
            const addSegment = async (el: HTMLElement) => {
                if (!el) return;
                const canvas = await html2canvas(el, {
                    scale: 2,
                    useCORS: true,
                    backgroundColor: null,
                    // Ignore anything marked no-pdf
                    ignoreElements: (node) => {
                        try {
                            const anyEl = node as HTMLElement;
                            return !!(anyEl?.dataset?.html2canvasIgnore) || anyEl?.classList?.contains('no-pdf');
                        } catch { return false; }
                    }
                });

                const imgData = canvas.toDataURL('image/png');
                const imgProps = pdf.getImageProperties(imgData);
                const imgWmm = contentWidth;
                const imgHmm = (imgWmm * imgProps.height) / imgProps.width;

                // If it doesn't fit on current page, start a new page
                if (cursorY + imgHmm > pageHeight - bottomMargin) {
                    pdf.addPage();
                    cursorY = topMargin;
                }

                pdf.addImage(imgData, 'PNG', margin, cursorY, imgWmm, imgHmm);
                cursorY += imgHmm;
            };

            // Grab the three segments
            const topEl = document.getElementById('PdfEmployeeInfoSegment') as HTMLElement | null;
            const itemsEl = document.getElementById('PdfItemsSegment') as HTMLElement | null; // contains ItemsSummaryStack now
            const itemsInstructions = document.getElementById('PdfInstructionsSegment') as HTMLElement | null;
            const bottomEl = document.getElementById('PdfApprovalsSegment') as HTMLElement | null;

            // Add in order; DO NOT force extra pages unless segment doesn't fit
            if (topEl) await addSegment(topEl);
            if (itemsEl) await addSegment(itemsEl);
            if (itemsInstructions) await addSegment(itemsInstructions);
            if (bottomEl) await addSegment(bottomEl);

            const safeEmp = (employeeName || 'employee').replace(/[^\w\s-]/g, '').trim() || 'employee';
            const ts = new Date().toISOString().slice(0, 10);
            const name = fileName || `PPE_Form_${safeEmp}_${ts}.pdf`;
            pdf.save(name);
        } catch (e: any) {
            onError?.('Failed to export PDF: ' + (e?.message || e));
        } finally {
            onExportModeChange(false);
            onBusyChange?.(false);
        }
    }, [onBusyChange, pdfMarginMm, onExportModeChange, onError]);

    return (
        <>
            {showButton && (
                <DefaultButton
                    text={exportMode ? 'Preparing…' : 'Export PDF'}
                    onClick={exportPdf}
                    disabled={disabled || exportMode || !isClosedBySystem}
                    style={{ marginLeft: 8 }}
                />
            )}
        </>
    );
};

export default ExportPdfControls;