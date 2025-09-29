
import * as React from "react";

type DocMetaProps = {
    docCode?: string;
    version?: string;
    effectiveDate?: string;
    page?: string | number;
};

export const DocumentMetaBanner: React.FC<DocMetaProps> = ({
    docCode = 'COR-HSE-01-FOR-001',
    version = 'V03',
    effectiveDate = '16-SEP-2020',
    page = 1
}) => {
    const grid: React.CSSProperties = {
        display: 'grid',
        gridTemplateColumns: '240px 1fr 260px 44px',
        border: '1px solid #000',
        margin: '12px 0',
        fontSize: 12,
        lineHeight: 1.3,
        background: '#fff'
    };
    const cell: React.CSSProperties = {
        borderRight: '1px solid #000',
        padding: '6px 8px',
        display: 'flex',
        alignItems: 'flex-start',
        justifyContent: 'center',
        flexDirection: 'column',
        gap: '6px',
        flexWrap: 'nowrap',
        // maxWidth: '65%'
    };
    return (
        <div style={grid}>
            <div style={{ ...cell, flexDirection: 'column', gap: 6 ,maxWidth: '60%'}}>
                <div>{docCode}</div>
                <div>Version: {version}</div>
            </div>
            <div style={{ ...cell, justifyContent: 'flex-start !important', alignItems: 'flex-start !important' }}>
                This document is confidential and property of The Coral Oil Co.
            </div>
            <div style={{ ...cell, justifyContent: 'center', alignItems: 'flex-end !important' }}>
                Effective Date: {effectiveDate}
            </div>
            <div style={{ ...cell, borderRight: 0, justifyContent: 'center', fontWeight: 600 }}>
                {page}
            </div>
        </div>
    );
};