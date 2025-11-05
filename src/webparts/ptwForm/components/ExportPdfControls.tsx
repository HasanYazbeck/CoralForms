import * as React from 'react';
import { DefaultButton } from '@fluentui/react';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';


export type ExportPdfControlsProps = {
    // The root element to capture (your form containerRef)
    targetRef: React.RefObject<HTMLElement>;
    // For filename
    originator?: string;
    fileName?: string;
    coralReferenceNumber?: string;
    // Parent-managed export mode toggle so PpeForm can enable controls while exporting
    exportMode: boolean;
    onExportModeChange: (on: boolean) => void;
    disabled?: boolean;
    showButton?: boolean;
    // Errors to host banner/UI
    onError?: (message: string) => void;
    onBusyChange?: (busy: boolean) => void; // NEW
    pdfMarginMm?: number;
};

const ExportPdfControls: React.FC<ExportPdfControlsProps> = ({
    targetRef,
    coralReferenceNumber,
    originator,
    fileName,
    exportMode,
    onExportModeChange,
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
            await new Promise(r => setTimeout(r, 50));
            if (!targetRef.current) { onError?.('Nothing to export.'); return; }

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
            const formTitleSection = document.getElementById('formTitleSection') as HTMLElement | null;
            const formHeader = document.getElementById('formHeaderInfo') as HTMLElement | null;
            const permitScheduleSectionContainer = document.getElementById('permitScheduleSectionContainer') as HTMLElement | null;
            const hacClassificationWorkAreaSection = document.getElementById('hacClassificationWorkAreaSection') as HTMLElement | null;
            const workHazardSection = document.getElementById('workHazardSection') as HTMLElement | null;
            const riskAssessmentListSection = document.getElementById('riskAssessmentListSection') as HTMLElement | null;
            const precautionsSection = document.getElementById('precautionsSection') as HTMLElement | null;
            const gasTestFireWatchAttachmentsSection = document.getElementById('gasTestFireWatchAttachmentsSection') as HTMLElement | null;
            const protectiveSafetyEquipmentSection = document.getElementById('protectiveSafetyEquipmentSection') as HTMLElement | null;
            const machineryToolsSection = document.getElementById('machineryToolsSection') as HTMLElement | null;
            const personnelInvolvedSection = document.getElementById('personnelInvolvedSection') as HTMLElement | null;
            const InstructionsSection = document.getElementById('InstructionsSection') as HTMLElement | null;
            const toolboxTalkSection = document.getElementById('toolboxTalkSection') as HTMLElement | null;
            const ptwSignOffSection = document.getElementById('ptwSignOffSection') as HTMLElement | null;
            const highRiskApprovalSection = document.getElementById('highRiskApprovalSection') as HTMLElement | null;


            // Add in order; DO NOT force extra pages unless segment doesn't fit
            if (formTitleSection) await addSegment(formTitleSection);
            if (formHeader) await addSegment(formHeader);
            if (permitScheduleSectionContainer) await addSegment(permitScheduleSectionContainer);
            if (hacClassificationWorkAreaSection) await addSegment(hacClassificationWorkAreaSection);
            if (workHazardSection) await addSegment(workHazardSection);
            if (riskAssessmentListSection) await addSegment(riskAssessmentListSection);
            if (precautionsSection) await addSegment(precautionsSection);
            if (gasTestFireWatchAttachmentsSection) await addSegment(gasTestFireWatchAttachmentsSection);
            if (protectiveSafetyEquipmentSection) await addSegment(protectiveSafetyEquipmentSection);
            if (machineryToolsSection) await addSegment(machineryToolsSection);
            if (personnelInvolvedSection) await addSegment(personnelInvolvedSection);
            if (InstructionsSection) await addSegment(InstructionsSection);
            if (toolboxTalkSection) await addSegment(toolboxTalkSection);
            if (ptwSignOffSection) await addSegment(ptwSignOffSection);
            if (highRiskApprovalSection) await addSegment(highRiskApprovalSection);
            // DONE - save the PDF
            const safeEmp = (originator || 'originator').replace(/[^\w\s-]/g, '').trim() || 'originator';
            const ts = new Date().toISOString().slice(0, 10);
            const name = fileName || `${coralReferenceNumber}_${safeEmp}_${ts}.pdf`;
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
                    disabled={disabled || exportMode}
                    style={{ marginLeft: 8 }}
                />
            )}
        </>
    );
};

export default ExportPdfControls;