import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { DefaultButton } from '@fluentui/react';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { DocumentMetaBanner } from '../../../Components/DocumentMetaBanner';


export type ExportPdfControlsProps = {
    // The root element to capture (your form containerRef)
    targetRef: React.RefObject<HTMLElement>;
    originator?: string;
    fileName?: string;
    coralReferenceNumber?: string;
    // Parent-managed export mode toggle so PpeForm can enable controls while exporting
    exportMode: boolean;
    onExportModeChange: (on: boolean) => void;
    disabled?: boolean;
    showButton?: boolean;
    onError?: (message: string) => void;
    onBusyChange?: (busy: boolean) => void; // NEW
    pdfMarginMm?: number;
    docCode?: string;
    companyName?: string;
    docVersion?: string;
    effectiveDate?: string;
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
    docCode,
    docVersion,
    effectiveDate,
    companyName
}) => {

    const exportPdf = React.useCallback(async () => {
        try {
            onBusyChange?.(true);
            onExportModeChange?.(true);
            await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
            if (!targetRef.current) { onError?.('Nothing to export.'); return; }

            const pdf = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' });
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = pdf.internal.pageSize.getHeight();
            const margin = pdfMarginMm;        // mm
            const topMargin = margin;          // you can customize differently if you want
            const bottomMargin = margin;
            const contentWidth = pageWidth - 2 * margin;

            // Helper: render footer banner (with page number) off-screen and return image + height in mm
            const renderFooter = async (pageNo: number) => {
                const host = document.createElement('div');
                host.style.position = 'fixed';
                host.style.left = '-10000px';
                host.style.top = '-10000px';
                host.style.background = '#ffffff';
                document.body.appendChild(host);

                ReactDOM.render(
                    React.createElement(DocumentMetaBanner, {
                        docCode: docCode,
                        version: docVersion,
                        effectiveDate: effectiveDate,
                        page: pageNo,
                        companyName: companyName || 'Coral'
                    }),
                    host
                );
                await new Promise(r => requestAnimationFrame(r));
                const canvas = await html2canvas(host, {
                    scale: 2,
                    useCORS: true,
                    backgroundColor: '#ffffff'
                });

                ReactDOM.unmountComponentAtNode(host);
                document.body.removeChild(host);

                const imgData = canvas.toDataURL('image/png');
                const hmm = (contentWidth * canvas.height) / canvas.width;
                return { imgData, hmm };
            };

            // Measure footer height (page 1 as sample)
            const sampleFooter = await renderFooter(1);
            const footerHmmSample = sampleFooter.hmm;
            const pageContentHeightMm = pageHeight - topMargin - bottomMargin - footerHmmSample;

            let cursorYmm = topMargin;
            let pageIndex = 1;

            const finishPage = async () => {
                // Draw footer for this page at the bottom
                const { imgData, hmm } = await renderFooter(pageIndex);
                const footerY = pageHeight - bottomMargin - hmm;
                pdf.addImage(imgData, 'PNG', margin, footerY, contentWidth, hmm);
                // Start new page
                pdf.addPage();
                pageIndex += 1;
                cursorYmm = topMargin;
            };

            // let cursorY = topMargin;
            const addSegment = async (el: HTMLElement) => {
                if (!el) return;
                const canvas = await html2canvas(el, {
                    scale: 2,
                    useCORS: true,
                    backgroundColor: '#ffffff',
                    // Ignore anything marked no-pdf
                    ignoreElements: (node) => {
                        try {
                            const anyEl = node as HTMLElement;
                            return !!(anyEl?.dataset?.html2canvasIgnore) || anyEl?.classList?.contains('no-pdf');
                        } catch { return false; }
                    }
                });
                const pxPerMm = canvas.width / contentWidth;
                const sectionHmm = canvas.height / pxPerMm;

                // If section fits entirely on a page
                if (sectionHmm <= pageContentHeightMm) {
                    // If it doesn't fit on the remaining space, start a new page first
                    const remainingMm = (topMargin + pageContentHeightMm) - cursorYmm;
                    if (remainingMm < sectionHmm - 0.01) {
                        await finishPage();
                    }

                    const imgData = canvas.toDataURL('image/jpeg', 0.95);
                    pdf.addImage(imgData, 'JPEG', margin, cursorYmm, contentWidth, sectionHmm);
                    cursorYmm += sectionHmm;
                    return;
                }

                let sYpx = 0;
                while (sYpx < canvas.height) {
                    const remainingMm = (topMargin + pageContentHeightMm) - cursorYmm;
                    if (remainingMm <= 0.01) {
                        await finishPage();
                    }

                    const remainingPx = Math.max(1, Math.floor(remainingMm * pxPerMm));
                    const sliceHpx = Math.min(remainingPx, canvas.height - sYpx);
                    const sliceHmm = sliceHpx / pxPerMm;

                    const pageCanvas = document.createElement('canvas');
                    pageCanvas.width = canvas.width;
                    pageCanvas.height = sliceHpx;

                    const ctx = pageCanvas.getContext('2d')!;
                    ctx.fillStyle = '#ffffff';
                    ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
                    ctx.drawImage(canvas, 0, sYpx, canvas.width, sliceHpx, 0, 0, canvas.width, sliceHpx);

                    const imgData = pageCanvas.toDataURL('image/jpeg', 0.95);
                    pdf.addImage(imgData, 'JPEG', margin, cursorYmm, contentWidth, sliceHmm);

                    cursorYmm += sliceHmm;
                    sYpx += sliceHpx;

                    // If exactly filled, the loop will trigger finishPage on next iteration
                }
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
            const urgentApprovalSection = document.getElementById('urgentApprovalSection') as HTMLElement | null;
            const ptwClosureSection = document.getElementById('ptwClosureSection') as HTMLElement | null;

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
            if (urgentApprovalSection) await addSegment(urgentApprovalSection);
            if (ptwClosureSection) await addSegment(ptwClosureSection);

            // Draw footer on the last page
            const lastFooter = await renderFooter(pageIndex);
            const lastFooterY = pageHeight - bottomMargin - lastFooter.hmm;
            pdf.addImage(lastFooter.imgData, 'PNG', margin, lastFooterY, contentWidth, lastFooter.hmm);

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
    }, [onBusyChange, pdfMarginMm, onExportModeChange, onError, docCode, docVersion, effectiveDate, originator, fileName, coralReferenceNumber, targetRef]);

    return (
        <>
            {showButton && (
                <DefaultButton
                    text={exportMode ? 'Preparing…' : 'Export to PDF'}
                    onClick={exportPdf}
                    disabled={disabled || exportMode}
                    style={{ marginLeft: 8 }}
                />
            )}
        </>
    );
};

export default ExportPdfControls;