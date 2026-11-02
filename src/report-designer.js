/*
 * Report Designer region plug-in for Oracle APEX (TechRevX, www.techrevx.com)
 *
 * techrevx.reportDesigner.init( regionId ) is called by the render procedure. It reads the settings and the template
 * list from #<regionId>_report_cfg, loads the region's data with AJAX (DATA) and builds the designer (or, in Report
 * mode, the template picker with preview and downloads) in #<regionId>_report.
 *
 * A template is JSON:
 *   { version: 1,
 *     page:     { size: "A4", orientation: "portrait", margin: 40 },                 sizes and margins in points
 *     style:    { fontSize: 10, accent: "#0f766e", text: "#1f2937" },
 *     settings: { groupBy: "", groupPageBreak: true, sheetPerGroup: true, dateFormat: "DD MMM YYYY", currency: "$",
 *                 fileName: "{REPORT}" },
 *     header:   { left: "", right: "", logo: null, line: false },
 *     footer:   { left: "{REPORT}", right: "Page {PAGE} of {PAGES}" },
 *     blocks:   [ { id, type: heading | text | fields | table | image | divider | spacer | pagebreak, ... } ] }
 * Texts take fields: {COLUMN}, {COLUMN|format}, {SUM:COLUMN|format} (SUM, AVG, MIN, MAX, COUNT), {COUNT}, {TODAY},
 * {NOW}, {REPORT}, {APP_USER}, {GROUP}, and in the header and footer {PAGE} and {PAGES}.
 *
 * With settings.groupBy the blocks repeat for each value of that column (one section per value, for example one
 * invoice per order); fields and totals then use the section's rows.
 *
 * Every output is built from the same intermediate model (build()), so the preview, PDF (pdfmake), Word (docx) and
 * Excel (ExcelJS) show the same report. The libraries are loaded on the first download.
 *
 * apex.region( id ).generate( format ) / refresh() / openTemplate( id ) / getTemplate() / setTemplate( json ) / save()
 * Events on the region: tx-report-generate { format, fileName, templateId, templateName }, tx-report-save { id, name }.
 */
( function( $, apex ) {
    "use strict";

    const ns = window.techrevx = window.techrevx || {};
    const esc = v => apex.util.escapeHTML( v === null || v === undefined ? "" : String( v ) );
    const uid = () => "b" + Date.now().toString( 36 ) + Math.random().toString( 36 ).substr( 2, 5 );
    const clone = o => JSON.parse( JSON.stringify( o ) );
    const abs = url => new URL( url, document.baseURI ).href;
    const PX = 96 / 72;   // CSS pixels per point

    // ---------------------------------------------------------------------------------------------------------------
    // libraries, loaded on the first download of each format
    // ---------------------------------------------------------------------------------------------------------------
    const LIBS = {
        pdf: { files: [ "pdfmake.min.js", "vfs_fonts.js" ], global: "pdfMake" },
        docx: { files: [ "docx.iife.js" ], global: "docx" },
        xlsx: { files: [ "exceljs.min.js" ], global: "ExcelJS" }
    };
    const libPromises = {};

    function loadScript( src ) {
        return new Promise( ( resolve, reject ) => {
            const s = document.createElement( "script" );
            // the libraries register as AMD modules when a module loader is on the page (Oracle JET): load them as globals
            const amd = window.define;
            window.define = undefined;
            const done = ok => () => {
                window.define = amd;
                if ( ok ) { resolve(); } else { reject( new Error( "Report Designer: could not load " + src ) ); }
            };
            s.src = src;
            s.onload = done( true );
            s.onerror = done( false );
            document.head.appendChild( s );
        } );
    }

    function loadLib( prefix, format ) {
        const lib = LIBS[ format ];
        if ( window[ lib.global ] && format !== "pdf" ) { return Promise.resolve( window[ lib.global ] ); }
        if ( !libPromises[ format ] ) {
            libPromises[ format ] = lib.files.reduce( ( p, f ) => p.then( () => loadScript( abs( prefix ) + f ) ), Promise.resolve() )
                .then( () => window[ lib.global ] )
                .catch( e => { delete libPromises[ format ]; throw e; } );
        }
        return libPromises[ format ];
    }

    // ---------------------------------------------------------------------------------------------------------------
    // template model
    // ---------------------------------------------------------------------------------------------------------------
    const PAGE_SIZES = { A4: [ 595.28, 841.89 ], Letter: [ 612, 792 ], Legal: [ 612, 1008 ], A3: [ 841.89, 1190.55 ], A5: [ 419.53, 595.28 ] };
    const EXCEL_PAPER = { Letter: 1, Legal: 5, A3: 8, A4: 9, A5: 11 };
    const DATE_FORMATS = [ "DD MMM YYYY", "D MMMM YYYY", "MMMM D, YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "MM/DD/YYYY", "DD.MM.YYYY" ];
    const FORMATS = [ [ "", "Automatic" ], [ "text", "Text" ], [ "number", "Number (1,234.50)" ], [ "integer", "Whole number (1,235)" ],
                      [ "currency", "Currency" ], [ "percent", "Percent (12.5%)" ], [ "date", "Date" ], [ "datetime", "Date and time" ] ];
    const TOTALS = [ [ "", "None" ], [ "sum", "Sum" ], [ "avg", "Average" ], [ "count", "Count" ], [ "min", "Minimum" ], [ "max", "Maximum" ] ];
    const TOTAL_NAMES = { sum: "Sum", avg: "Average", count: "Count", min: "Minimum", max: "Maximum" };
    const WIDTHS = [ [ 0.6, "Narrow" ], [ 1, "Normal" ], [ 1.6, "Wide" ], [ 2.4, "Extra wide" ] ];
    const HEADING_SIZES = { 1: 22, 2: 15, 3: 12 };
    const MUTED = "#6b7280", LINE = "#d9dde3", ZEBRA = "#f5f6f8";

    const labelOf = name => {
        const s = String( name || "" ).toLowerCase().replace( /_+/g, " " ).trim();
        return s.charAt( 0 ).toUpperCase() + s.slice( 1 );
    };
    const isKey = name => /(^|_)(ID|NO|NUM|NUMBER|CODE|YEAR|LINE|SEQ|ZIP)$/i.test( name );
    const isMoney = name => /(AMOUNT|PRICE|TOTAL|COST|SALARY|REVENUE|SALES|VALUE|BALANCE|FEE|PAY)/i.test( name );
    // amounts add up; prices and rates do not
    const isSummable = name => isMoney( name ) && !/(PRICE|RATE)/i.test( name );

    function guessFormat( c ) {
        if ( c.type === "date" ) { return "date"; }
        if ( c.type !== "number" ) { return ""; }
        if ( isKey( c.name ) ) { return "text"; }
        return isMoney( c.name ) ? "number" : "";
    }

    function colDef( c ) {
        const num = c.type === "number" && !isKey( c.name );
        return { field: c.name, label: labelOf( c.name ), align: num ? "right" : "left", format: guessFormat( c ),
                 total: num && isSummable( c.name ) ? "sum" : "", width: 1 };
    }

    const BLOCKS = {
        heading: { label: "Heading", icon: "fa-header", hint: "Title or section heading",
                   make: () => ( { text: "Heading", level: 2, align: "left", color: "" } ) },
        text: { label: "Text", icon: "fa-align-left", hint: "Paragraph with fields",
                make: () => ( { text: "Type your text here. Insert fields with the { } button.", size: 0, bold: false, italic: false, align: "left", color: "" } ) },
        fields: { label: "Fields", icon: "fa-list-alt", hint: "Labels with values",
                  make: d => ( { columns: 2, align: "left",
                                 items: d.columns.slice( 0, 4 ).map( c => ( { label: labelOf( c.name ), value: "{" + c.name + "}" } ) ) } ) },
        table: { label: "Table", icon: "fa-table", hint: "Rows of data with totals",
                 make: d => ( { columns: d.columns.slice( 0, 6 ).map( colDef ), striped: true, border: "lines", totals: true,
                                groupBy: "", subtotals: true, size: 0 } ) },
        image: { label: "Image", icon: "fa-image", hint: "Logo or picture",
                 make: () => ( { src: "", ratio: 0.5, width: 140, align: "left" } ) },
        divider: { label: "Line", icon: "fa-minus", hint: "Horizontal line",
                   make: () => ( { color: "", thickness: 1 } ) },
        spacer: { label: "Space", icon: "fa-arrows-v", hint: "Empty space",
                  make: () => ( { height: 16 } ) },
        pagebreak: { label: "Page break", icon: "fa-scissors", hint: "Next block on a new page",
                     make: () => ( {} ) }
    };

    function defaults() {
        return {
            version: 1,
            page: { size: "A4", orientation: "portrait", margin: 40 },
            style: { fontSize: 10, accent: "#0f766e", text: "#1f2937" },
            settings: { groupBy: "", groupPageBreak: true, sheetPerGroup: true, dateFormat: "DD MMM YYYY", currency: "$", fileName: "{REPORT}" },
            header: { left: "", right: "", logo: null, line: false },
            footer: { left: "{REPORT}", right: "Page {PAGE} of {PAGES}" },
            blocks: []
        };
    }

    function normalize( t ) {
        const d = defaults(), src = t && typeof t === "object" ? t : {}, r = Object.assign( {}, d, src );
        [ "page", "style", "settings", "header", "footer" ].forEach( k => { r[ k ] = Object.assign( {}, d[ k ], src[ k ] ); } );
        r.blocks = Array.isArray( src.blocks )
            ? src.blocks.filter( b => b && BLOCKS[ b.type ] ).map( b => Object.assign( { id: uid() }, b ) )
            : [];
        return r;
    }

    function pageSize( t ) {
        const s = PAGE_SIZES[ t.page.size ] || PAGE_SIZES.A4;
        return t.page.orientation === "landscape" ? [ s[ 1 ], s[ 0 ] ] : [ s[ 0 ], s[ 1 ] ];
    }

    // starting points offered by New template, made from the region's columns
    function starter( kind, data, groupCol ) {
        const t = normalize( null ), cols = data.columns;
        const add = ( type, props ) => { const b = Object.assign( { id: uid(), type }, BLOCKS[ type ].make( data ), props || {} ); t.blocks.push( b ); return b; };
        const money = cols.filter( c => c.type === "number" && !isKey( c.name ) && isSummable( c.name ) );
        const sumField = money.length ? money[ money.length - 1 ].name : null;

        if ( kind === "blank" ) {
            add( "heading", { text: "{REPORT}", level: 1 } );
            return t;
        }
        if ( kind === "list" || kind === "grouped" ) {
            add( "heading", { text: "{REPORT}", level: 1 } );
            add( "text", { text: "Created on {TODAY}, {COUNT} rows.", size: 9, color: MUTED } );
            const tb = add( "table", { columns: cols.filter( c => c.name !== ( kind === "grouped" ? groupCol : "" ) ).slice( 0, 8 ).map( colDef ) } );
            if ( kind === "grouped" && groupCol ) { tb.groupBy = groupCol; tb.subtotals = true; }
            return t;
        }
        // document: one section per value of groupCol, on its own page (for example one invoice per order)
        const groups = groupRows( data.rows, groupCol );
        const constant = cols.filter( c => c.name !== groupCol && groups.length && groups.every( g => g.rows.every( r => same( r[ c.name ], g.rows[ 0 ][ c.name ] ) ) ) );
        const varying = cols.filter( c => c.name !== groupCol && constant.indexOf( c ) < 0 );
        t.settings.groupBy = groupCol || "";
        t.settings.groupPageBreak = true;
        t.settings.fileName = "{REPORT}";
        t.header.right = "{TODAY}";
        t.header.line = true;
        add( "heading", { text: groupCol ? labelOf( groupCol ) + " {" + groupCol + "}" : "{REPORT}", level: 1 } );
        if ( constant.length ) {
            add( "fields", { columns: 2, items: constant.slice( 0, 8 ).map( c => ( { label: labelOf( c.name ), value: "{" + c.name + "}" } ) ) } );
        }
        add( "spacer", { height: 10 } );
        // the total goes below the table, as on an invoice
        add( "table", { columns: ( varying.length ? varying : cols ).slice( 0, 7 ).map( colDef ), totals: !sumField } );
        if ( sumField ) {
            add( "fields", { columns: 1, align: "right", items: [ { label: "Total", value: "{SUM:" + sumField + "|number}" } ] } );
        }
        return t;
    }

    // ---------------------------------------------------------------------------------------------------------------
    // data, groups and formatting
    // ---------------------------------------------------------------------------------------------------------------
    function toDate( s ) {
        const m = /^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d):(\d\d))?/.exec( s || "" );
        return m ? new Date( +m[ 1 ], m[ 2 ] - 1, +m[ 3 ], +( m[ 4 ] || 0 ), +( m[ 5 ] || 0 ), +( m[ 6 ] || 0 ) ) : null;
    }
    const keyOf = v => v instanceof Date ? v.getTime() : v;
    const same = ( a, b ) => keyOf( a ) === keyOf( b );

    class Data {
        constructor( res ) {
            res = res || {};
            this.columns = ( res.columns || [] ).map( c => ( { name: String( c.name ).toUpperCase(), type: c.type } ) );
            this.types = {};
            this.columns.forEach( c => { this.types[ c.name ] = c.type; } );
            this.rows = ( res.rows || [] ).map( r => {
                const o = {};
                this.columns.forEach( ( c, i ) => { o[ c.name ] = c.type === "date" && r[ i ] ? toDate( r[ i ] ) : r[ i ]; } );
                return o;
            } );
            this.truncated = !!res.truncated;
        }
        has( name ) { return Object.prototype.hasOwnProperty.call( this.types, name ); }
    }

    function groupRows( rows, col ) {
        if ( !col ) { return [ { key: null, rows } ]; }
        const map = new Map();
        rows.forEach( r => {
            const k = keyOf( r[ col ] );
            if ( !map.has( k ) ) { map.set( k, { key: r[ col ], rows: [] } ); }
            map.get( k ).rows.push( r );
        } );
        return map.size ? [ ...map.values() ] : [ { key: null, rows: [] } ];
    }

    function aggregate( rows, field, fn ) {
        const vals = rows.map( r => r[ field ] ).filter( v => v !== null && v !== undefined && v !== "" );
        if ( fn === "count" ) { return vals.length; }
        const dates = vals.length && vals.every( v => v instanceof Date );
        const nums = vals.map( v => v instanceof Date ? v.getTime() : Number( v ) ).filter( v => !isNaN( v ) );
        if ( !nums.length ) { return null; }
        switch ( fn ) {
            case "sum": return dates ? null : nums.reduce( ( a, b ) => a + b, 0 );
            case "avg": return dates ? null : nums.reduce( ( a, b ) => a + b, 0 ) / nums.length;
            // the earliest and latest of dates stay dates
            case "min": return dates ? new Date( Math.min( ...nums ) ) : Math.min( ...nums );
            case "max": return dates ? new Date( Math.max( ...nums ) ) : Math.max( ...nums );
        }
        return null;
    }

    class Formatter {
        constructor( tpl ) {
            this.s = tpl.settings;
            this.locale = ( apex.locale && apex.locale.getLanguage && apex.locale.getLanguage() ) || undefined;
            try { new Intl.NumberFormat( this.locale ); } catch ( e ) { this.locale = undefined; }
            this.cache = {};
        }
        num( v, min, max, grouping ) {
            const k = min + "," + max + "," + grouping;
            if ( !this.cache[ k ] ) {
                this.cache[ k ] = new Intl.NumberFormat( this.locale, { minimumFractionDigits: min, maximumFractionDigits: max, useGrouping: grouping } );
            }
            return this.cache[ k ].format( v );
        }
        date( d, time ) {
            const p = n => String( n ).padStart( 2, "0" );
            const mon = len => new Intl.DateTimeFormat( this.locale, { month: len } ).format( d );
            const s = ( this.s.dateFormat || "DD MMM YYYY" ).replace( /YYYY|MMMM|MMM|MM|DD|D/g, t => String( {
                YYYY: d.getFullYear(), MMMM: mon( "long" ), MMM: mon( "short" ), MM: p( d.getMonth() + 1 ), DD: p( d.getDate() ), D: d.getDate()
            }[ t ] ) );
            return time ? s + " " + p( d.getHours() ) + ":" + p( d.getMinutes() ) : s;
        }
        value( v, type, format ) {
            if ( v === null || v === undefined || v === "" ) { return ""; }
            const f = format || "";
            if ( v instanceof Date ) { return this.date( v, f === "datetime" ); }
            if ( typeof v === "number" ) {
                switch ( f ) {
                    case "text": return String( v );
                    case "number": return this.num( v, 2, 2, true );
                    case "integer": return this.num( Math.round( v ), 0, 0, true );
                    case "currency": return ( v < 0 ? "-" : "" ) + ( this.s.currency || "" ) + this.num( Math.abs( v ), 2, 2, true );
                    case "percent": return this.num( v, 0, 1, true ) + "%";
                    case "date": case "datetime": return String( v );
                    default: return this.num( v, 0, 2, false );
                }
            }
            if ( type === "number" && f && f !== "text" && !isNaN( Number( v ) ) ) { return this.value( Number( v ), type, f ); }
            return String( v );
        }
    }

    // ---------------------------------------------------------------------------------------------------------------
    // fields in texts: {COLUMN}, {COLUMN|format}, {SUM:COLUMN|format}, {COUNT}, {TODAY}, ...
    // ---------------------------------------------------------------------------------------------------------------
    const BUILTINS = [ "TODAY", "NOW", "REPORT", "APP_USER", "GROUP", "COUNT", "PAGE", "PAGES" ];

    function parts( text ) {
        const out = [], re = /\{([^{}\n]{1,80})\}/g;
        let last = 0, m;
        const s = String( text === null || text === undefined ? "" : text );
        while ( ( m = re.exec( s ) ) ) {
            const inner = m[ 1 ].trim(), bar = inner.indexOf( "|" );
            const expr = ( bar < 0 ? inner : inner.slice( 0, bar ) ).trim().toUpperCase();
            const fmt = bar < 0 ? "" : inner.slice( bar + 1 ).trim().toLowerCase();
            const colon = expr.indexOf( ":" ), name = ( colon < 0 ? expr : expr.slice( colon + 1 ) ).trim();
            // only names count as fields: "{ }" or "{1,2}" stay as written
            if ( !/^[A-Z_][A-Z0-9_$#]*$/.test( name ) ) { continue; }
            if ( m.index > last ) { out.push( { text: s.slice( last, m.index ) } ); }
            out.push( { raw: m[ 0 ], fn: colon < 0 ? "" : expr.slice( 0, colon ).trim().toLowerCase(), name, fmt } );
            last = re.lastIndex;
        }
        if ( last < s.length ) { out.push( { text: s.slice( last ) } ); }
        return out;
    }

    // value of one field: { ok, text } (PAGE and PAGES stay as they are; each output fills them in)
    function evalToken( p, ctx ) {
        const f = ctx.f, d = ctx.data;
        if ( p.fn ) {
            if ( !TOTAL_NAMES[ p.fn ] || !d.has( p.name ) ) { return { ok: false }; }
            const v = aggregate( ctx.rows, p.name, p.fn );
            const type = p.fn === "count" ? "number" : d.types[ p.name ];
            return { ok: true, text: f.value( v, type, p.fmt || ( p.fn === "count" ? "integer" : p.fn === "avg" && type === "number" ? "number" : "" ) ) };
        }
        if ( d.has( p.name ) ) { return { ok: true, text: f.value( ctx.row[ p.name ], d.types[ p.name ], p.fmt ) }; }
        switch ( p.name ) {
            case "TODAY": return { ok: true, text: f.date( new Date() ) };
            case "NOW": return { ok: true, text: f.date( new Date(), true ) };
            case "REPORT": return { ok: true, text: ctx.name || "Report" };
            case "APP_USER": return { ok: true, text: ctx.user || "" };
            case "GROUP": return { ok: true, text: ctx.group === null || ctx.group === undefined ? "" : f.value( ctx.group, d.types[ ctx.groupBy ], p.fmt ) };
            case "COUNT": return { ok: true, text: f.value( ctx.rows.length, "number", p.fmt || "integer" ) };
            case "PAGE": return { ok: true, page: true, text: p.raw };
            case "PAGES": return { ok: true, page: true, text: p.raw };
        }
        return { ok: false };
    }

    function resolveText( text, ctx ) {
        return parts( text ).map( p => p.raw === undefined ? p.text : ( r => r.ok ? r.text : p.raw )( evalToken( p, ctx ) ) ).join( "" );
    }

    // preview: values (or the field names) marked up, unknown fields flagged
    function resolveHtml( text, ctx, names, page ) {
        return parts( text ).map( p => {
            if ( p.raw === undefined ) { return esc( p.text ); }
            const r = evalToken( p, ctx );
            if ( !r.ok ) { return "<span class=\"tx-Rd-tok is-bad\" title=\"Unknown field\">" + esc( p.raw ) + "</span>"; }
            if ( names ) { return "<span class=\"tx-Rd-tok\">" + esc( p.raw ) + "</span>"; }
            if ( r.page ) { return "<span class=\"tx-Rd-val\">" + esc( p.name === "PAGE" ? page || 1 : page || 1 ) + "</span>"; }
            return "<span class=\"tx-Rd-val\">" + esc( r.text ) + "</span>";
        } ).join( "" );
    }

    const pageTokens = ( s, cur, count ) => String( s || "" ).replace( /\{\s*PAGES\s*\}/gi, count ).replace( /\{\s*PAGE\s*\}/gi, cur );

    // ---------------------------------------------------------------------------------------------------------------
    // the intermediate model: every group with its blocks resolved
    // ---------------------------------------------------------------------------------------------------------------
    function build( tpl, data, opts ) {
        const f = new Formatter( tpl ), groupBy = data.has( tpl.settings.groupBy ) ? tpl.settings.groupBy : "";
        const groups = groupRows( data.rows, groupBy );
        const list = opts.maxGroups ? groups.slice( 0, opts.maxGroups ) : groups;
        const model = { tpl, f, groupBy, moreGroups: groups.length - list.length, groups: [] };
        list.forEach( g => {
            const ctx = { f, data, rows: g.rows, row: g.rows[ 0 ] || {}, group: g.key, groupBy, name: opts.name, user: opts.user };
            model.groups.push( { key: g.key, ctx, els: tpl.blocks.map( b => element( b, ctx, opts ) ) } );
        } );
        if ( !model.groups.length ) {
            const ctx = { f, data, rows: [], row: {}, group: null, groupBy, name: opts.name, user: opts.user };
            model.groups.push( { key: null, ctx, els: tpl.blocks.map( b => element( b, ctx, opts ) ) } );
        }
        model.doc = model.groups[ 0 ].ctx;
        return model;
    }

    function element( b, ctx, opts ) {
        const tx = s => resolveText( s, ctx ), el = { id: b.id, type: b.type, b };
        switch ( b.type ) {
            case "heading": return Object.assign( el, { text: tx( b.text ) } );
            case "text": return Object.assign( el, { text: tx( b.text ) } );
            case "fields": return Object.assign( el, { items: ( b.items || [] ).map( i => ( { label: tx( i.label ), value: tx( i.value ) } ) ) } );
            case "table": return Object.assign( el, tableModel( b, ctx, opts.limit ) );
        }
        return el;
    }

    // table rows: { kind: data | group | subtotal | total, cells: [ { v, type, format, text, align, total } ], text }
    function tableModel( b, ctx, limit ) {
        const d = ctx.data, f = ctx.f;
        const cols = ( b.columns || [] ).filter( c => c.field && d.has( c.field ) );
        const head = cols.map( c => ( { text: resolveText( c.label, ctx ), align: c.align || "left", width: +c.width || 1 } ) );
        const hasTotals = cols.some( c => c.total );
        const dataCell = ( r, c ) => {
            const type = d.types[ c.field ];
            return { v: r[ c.field ], type, format: c.format || ( type === "date" ? "date" : "" ), text: f.value( r[ c.field ], type, c.format ), align: c.align || "left" };
        };
        const totalRow = ( rows, kind, label ) => ( {
            kind, rows: rows.length,
            cells: cols.map( ( c, i ) => {
                if ( !c.total ) { return { text: i === 0 ? label : "", align: c.align || "left" }; }
                const v = aggregate( rows, c.field, c.total ), type = c.total === "count" ? "number" : d.types[ c.field ];
                const format = c.total === "count" ? "integer" : c.format || ( c.total === "avg" && type === "number" ? "number" : "" );
                return { v, type, format, text: f.value( v, type, format ), align: c.align || "left", total: c.total };
            } )
        } );
        const rows = [];
        let shown = 0, hidden = 0;
        const pushData = list => list.forEach( r => {
            if ( limit && shown >= limit ) { hidden++; return; }
            shown++;
            rows.push( { kind: "data", cells: cols.map( c => dataCell( r, c ) ) } );
        } );
        if ( b.groupBy && d.has( b.groupBy ) ) {
            groupRows( ctx.rows, b.groupBy ).forEach( g => {
                if ( limit && shown >= limit ) { hidden += g.rows.length; return; }
                rows.push( { kind: "group", text: labelOf( b.groupBy ) + ": " + f.value( g.key, d.types[ b.groupBy ], "" ), count: g.rows.length } );
                pushData( g.rows );
                if ( b.subtotals !== false && hasTotals ) { rows.push( totalRow( g.rows, "subtotal", "Subtotal" ) ); }
            } );
        } else {
            pushData( ctx.rows );
        }
        return { cols, head, rows, hidden, total: b.totals !== false && hasTotals ? totalRow( ctx.rows, "total", "Total" ) : null };
    }

    // ---------------------------------------------------------------------------------------------------------------
    // colors and images
    // ---------------------------------------------------------------------------------------------------------------
    function hex6( c, fallback ) {
        const m = /^#?([0-9a-f]{6}|[0-9a-f]{3})$/i.exec( String( c || "" ).trim() );
        if ( !m ) { return fallback === undefined ? "1F2937" : fallback; }
        const h = m[ 1 ].length === 3 ? m[ 1 ].replace( /./g, "$&$&" ) : m[ 1 ];
        return h.toUpperCase();
    }
    function mix( c, withColor, amount ) {
        const a = hex6( c ), b = hex6( withColor );
        return [ 0, 2, 4 ].map( i => Math.round( parseInt( a.substr( i, 2 ), 16 ) * ( 1 - amount ) + parseInt( b.substr( i, 2 ), 16 ) * amount )
            .toString( 16 ).padStart( 2, "0" ) ).join( "" ).toUpperCase();
    }
    const imageType = src => /^data:image\/png/i.test( src ) ? "png" : "jpg";
    const imageBytes = src => Uint8Array.from( atob( String( src ).split( "," )[ 1 ] || "" ), c => c.charCodeAt( 0 ) );

    // reads an image file into a PNG or JPEG data URL, at most 1200 px wide
    function readImage( file ) {
        return new Promise( ( resolve, reject ) => {
            const reader = new FileReader();
            reader.onerror = () => reject( new Error( "The image could not be read." ) );
            reader.onload = () => {
                const img = new Image();
                img.onerror = () => reject( new Error( "This file is not an image the browser can show." ) );
                img.onload = () => {
                    const w0 = img.naturalWidth || 300, h0 = img.naturalHeight || 150, scale = Math.min( 1, 1200 / w0 );
                    const keep = /^data:image\/(png|jpeg)/i.test( reader.result ) && scale === 1;
                    let src = reader.result;
                    if ( !keep ) {
                        const cv = document.createElement( "canvas" );
                        cv.width = Math.round( w0 * scale );
                        cv.height = Math.round( h0 * scale );
                        cv.getContext( "2d" ).drawImage( img, 0, 0, cv.width, cv.height );
                        src = cv.toDataURL( "image/png" );
                    }
                    resolve( { src, ratio: h0 / w0 } );
                };
                img.src = reader.result;
            };
            reader.readAsDataURL( file );
        } );
    }

    function saveBlob( blob, name ) {
        const a = document.createElement( "a" ), url = URL.createObjectURL( blob );
        a.href = url;
        a.download = name;
        document.body.appendChild( a );
        a.click();
        a.remove();
        setTimeout( () => URL.revokeObjectURL( url ), 60000 );
    }

    const widthsOf = ( weights, total ) => {
        const sum = weights.reduce( ( a, b ) => a + b, 0 ) || 1;
        return weights.map( w => total * w / sum );
    };

    // ---------------------------------------------------------------------------------------------------------------
    // PDF (pdfmake)
    // ---------------------------------------------------------------------------------------------------------------
    function toPdf( model, pdfMake ) {
        const t = model.tpl, [ W, H ] = pageSize( t ), m = +t.page.margin || 40, fs = +t.style.fontSize || 10;
        const accent = "#" + hex6( t.style.accent, "0F766E" ), text = "#" + hex6( t.style.text );
        const avail = W - 2 * m, h = t.header, ft = t.footer, doc = model.doc;
        const hasHeader = !!( h.left || h.right || h.logo ), hasFooter = !!( ft.left || ft.right );
        const headerH = hasHeader ? ( h.logo ? 44 : 22 ) + ( h.line ? 8 : 0 ) : 0, footerH = hasFooter ? 20 : 0;
        const content = [];

        model.groups.forEach( ( g, gi ) => {
            if ( gi > 0 ) { content.push( t.settings.groupPageBreak ? { text: "", pageBreak: "before" } : { text: "", margin: [ 0, 14, 0, 0 ] } ); }
            g.els.forEach( el => { const n = pdfNode( el, t, { avail, fs, accent, text } ); if ( n ) { content.push( n ); } } );
        } );
        // the space after the last block would start an empty page when the page is full
        const lastNode = content[ content.length - 1 ];
        if ( lastNode && lastNode.margin ) { lastNode.margin = lastNode.margin.slice( 0, 3 ).concat( 0 ); }

        const dd = {
            pageSize: { width: W, height: H },
            pageMargins: [ m, m + headerH, m, m + footerH ],
            content,
            defaultStyle: { fontSize: fs, color: text, lineHeight: 1.15 },
            // a heading near the bottom of a page moves to the next page, with what follows it
            pageBreakBefore: node => !!node.headlineLevel && node.startPosition && node.startPosition.verticalRatio > 0.82,
            info: { title: doc.name || "Report", creator: "Report Designer by TechRevX" }
        };
        if ( hasHeader ) {
            const left = resolveText( h.left, doc ), right = resolveText( h.right, doc );
            dd.header = ( cur, count ) => {
                const cols = [];
                if ( h.logo ) { cols.push( { image: h.logo, fit: [ 140, 36 ], width: "auto" } ); }
                cols.push( { text: pageTokens( left, cur, count ), width: "*", fontSize: fs, bold: true, margin: [ 0, h.logo ? 10 : 0, 0, 0 ] } );
                cols.push( { text: pageTokens( right, cur, count ), width: "auto", fontSize: fs - 1, color: MUTED, alignment: "right", margin: [ 0, h.logo ? 11 : 1, 0, 0 ] } );
                const stack = [ { columns: cols, columnGap: 12 } ];
                if ( h.line ) { stack.push( { canvas: [ { type: "line", x1: 0, y1: 6, x2: avail, y2: 6, lineWidth: 0.75, lineColor: LINE } ] } ); }
                return { margin: [ m, Math.max( 14, m * 0.55 ), m, 0 ], stack };
            };
        }
        if ( hasFooter ) {
            const left = resolveText( ft.left, doc ), right = resolveText( ft.right, doc );
            dd.footer = ( cur, count ) => ( {
                margin: [ m, 6, m, 0 ], fontSize: fs - 2, color: MUTED,
                columns: [ { text: pageTokens( left, cur, count ), width: "*" }, { text: pageTokens( right, cur, count ), width: "auto", alignment: "right" } ]
            } );
        }
        return pdfMake.createPdf( dd ).getBlob();
    }

    function pdfNode( el, t, c ) {
        const b = el.b;
        switch ( el.type ) {
            case "heading": {
                const lvl = +b.level || 2;
                return { text: el.text, headlineLevel: lvl, fontSize: HEADING_SIZES[ lvl ] || 15, bold: true, alignment: b.align || "left",
                         color: b.color ? "#" + hex6( b.color ) : lvl === 1 ? c.text : c.accent, margin: [ 0, lvl === 1 ? 0 : 8, 0, lvl === 1 ? 8 : 5 ] };
            }
            case "text":
                return { text: el.text, fontSize: +b.size || c.fs, bold: !!b.bold, italics: !!b.italic, alignment: b.align || "left",
                         color: b.color ? "#" + hex6( b.color ) : c.text, margin: [ 0, 0, 0, 6 ] };
            case "fields": {
                const n = Math.max( 1, Math.min( 3, +b.columns || 2 ) ), body = [];
                for ( let i = 0; i < el.items.length; i += n ) {
                    const row = [];
                    for ( let k = 0; k < n; k++ ) {
                        const it = el.items[ i + k ];
                        // right-aligned values end where a table's last column ends (its 4 pt padding)
                        row.push( { text: it ? it.label : "", color: MUTED, fontSize: c.fs - 0.5, margin: [ 0, 1, 6, 1 ] },
                                  { text: it ? it.value : "", bold: true, margin: [ 0, 1, b.align === "right" ? 0 : 16, 1 ], alignment: b.align === "right" ? "right" : "left" } );
                    }
                    body.push( row );
                }
                if ( !body.length ) { return null; }
                const widths = [];
                for ( let k = 0; k < n; k++ ) { widths.push( "auto", b.align === "right" ? "auto" : "*" ); }
                const table = { table: { widths, body }, layout: { defaultBorder: false, paddingLeft: () => 0, paddingRight: i => i % 2 ? 4 : 0, paddingTop: () => 0, paddingBottom: () => 0 },
                                margin: [ 0, 2, 0, 8 ] };
                return b.align === "right" ? { columns: [ { text: "", width: "*" }, Object.assign( table, { width: "auto" } ) ], margin: [ 0, 2, 0, 8 ] } : table;
            }
            case "table": {
                if ( !el.cols.length ) { return null; }
                const n = el.cols.length, size = +b.size || c.fs - 0.5, grid = b.border === "grid", none = b.border === "none";
                // full width: the page width less the cell padding (and the vertical lines of a grid)
                const widths = widthsOf( el.head.map( x => x.width ), c.avail - n * 8 - ( grid ? ( n + 1 ) * 0.5 : 0 ) );
                const tint = "#" + mix( c.accent, "#ffffff", 0.88 );
                const body = [ el.head.map( x => ( { text: x.text, bold: true, color: "#ffffff", fillColor: c.accent, alignment: x.align, fontSize: size } ) ) ];
                let zebra = 0;
                const totalCells = ( r, fill, top ) => r.cells.map( x => ( { text: x.text, bold: true, alignment: x.align, fontSize: size, fillColor: fill,
                                                                             border: top ? [ false, true, false, false ] : undefined } ) );
                el.rows.forEach( r => {
                    if ( r.kind === "group" ) {
                        zebra = 0;
                        const row = [ { text: r.text, colSpan: n, bold: true, fillColor: tint, color: c.accent, fontSize: size } ];
                        for ( let i = 1; i < n; i++ ) { row.push( {} ); }
                        body.push( row );
                    } else if ( r.kind === "subtotal" ) {
                        body.push( totalCells( r, null ) );
                    } else {
                        const fill = b.striped !== false && zebra++ % 2 === 1 ? ZEBRA : null;
                        body.push( r.cells.map( x => ( { text: x.text, alignment: x.align, fontSize: size, fillColor: fill } ) ) );
                    }
                } );
                if ( el.total ) { body.push( totalCells( el.total, tint ) ); }
                return {
                    table: { headerRows: 1, widths, body, dontBreakRows: true, keepWithHeaderRows: 1 },
                    layout: {
                        hLineWidth: ( i, node ) => none ? 0 : grid ? 0.5 : ( i === 0 || i === 1 || i === node.table.body.length ? 0 : 0.5 ),
                        vLineWidth: ( i, node ) => grid ? 0.5 : 0,
                        hLineColor: () => LINE, vLineColor: () => LINE,
                        paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 3.5, paddingBottom: () => 3.5
                    },
                    margin: [ 0, 2, 0, 10 ]
                };
            }
            case "image":
                return b.src ? { image: b.src, width: Math.min( +b.width || 140, c.avail ), alignment: b.align || "left", margin: [ 0, 2, 0, 8 ] } : null;
            case "divider":
                return { canvas: [ { type: "line", x1: 0, y1: 0, x2: c.avail, y2: 0, lineWidth: +b.thickness || 1, lineColor: b.color ? "#" + hex6( b.color ) : LINE } ],
                         margin: [ 0, 6, 0, 10 ] };
            case "spacer":
                return { text: "", margin: [ 0, Math.max( 0, ( +b.height || 16 ) - c.fs ), 0, 0 ] };
            case "pagebreak":
                return { text: "", pageBreak: "before" };
        }
        return null;
    }

    // ---------------------------------------------------------------------------------------------------------------
    // Word (docx)
    // ---------------------------------------------------------------------------------------------------------------
    function toDocx( model, D ) {
        const t = model.tpl, size = PAGE_SIZES[ t.page.size ] || PAGE_SIZES.A4, landscape = t.page.orientation === "landscape";
        const [ W ] = pageSize( t ), m = +t.page.margin || 40, fs = +t.style.fontSize || 10;
        const tw = pt => Math.round( pt * 20 );
        const c = { D, tw, fs, avail: W - 2 * m, accent: hex6( t.style.accent, "0F766E" ), text: hex6( t.style.text ) };
        const children = [];
        model.groups.forEach( ( g, gi ) => {
            if ( gi > 0 ) {
                children.push( t.settings.groupPageBreak ? new D.Paragraph( { children: [ new D.PageBreak() ] } )
                                                         : new D.Paragraph( { spacing: { before: 240 } } ) );
            }
            g.els.forEach( el => { const n = docxNode( el, c ); if ( n ) { children.push( ...[].concat( n ) ); } } );
        } );
        if ( !children.length ) { children.push( new D.Paragraph( "" ) ); }

        const h = t.header, ft = t.footer, doc = model.doc, none = { style: D.BorderStyle.NONE, size: 0, color: "FFFFFF" };
        const noBorders = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };
        // header and footer: a borderless two-column table (left text, right text) with page number fields
        const runsWithPages = ( s, o ) => parts( s ).map( p => {
            if ( p.raw !== undefined && !p.fn && ( p.name === "PAGE" || p.name === "PAGES" ) ) {
                return new D.TextRun( Object.assign( { children: [ p.name === "PAGE" ? D.PageNumber.CURRENT : D.PageNumber.TOTAL_PAGES ] }, o ) );
            }
            return new D.TextRun( Object.assign( { text: p.raw !== undefined ? resolveText( p.raw, doc ) : p.text }, o ) );
        } );
        const bar = ( left, right, o, logo ) => new D.Table( {
            width: { size: tw( c.avail ), type: D.WidthType.DXA },
            columnWidths: [ tw( c.avail * 0.6 ), tw( c.avail * 0.4 ) ],
            borders: noBorders,
            rows: [ new D.TableRow( { children: [
                new D.TableCell( { borders: noBorders, verticalAlign: "center", children: [
                    new D.Paragraph( { children: ( logo ? [ new D.ImageRun( { type: imageType( logo ), data: imageBytes( logo ),
                        transformation: imgBox( logo, 140, 36, t ) } ), new D.TextRun( "   " ) ] : [] ).concat( runsWithPages( left, o.left ) ) } ) ] } ),
                new D.TableCell( { borders: noBorders, verticalAlign: "center", children: [
                    new D.Paragraph( { alignment: D.AlignmentType.RIGHT, children: runsWithPages( right, o.right ) } ) ] } )
            ] } ) ]
        } );
        const section = {
            properties: { page: {
                size: { width: tw( size[ 0 ] ), height: tw( size[ 1 ] ), orientation: landscape ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT },
                margin: { top: tw( m + ( h.left || h.right || h.logo ? 26 : 0 ) ), bottom: tw( m + 10 ), left: tw( m ), right: tw( m ),
                          header: tw( Math.max( 14, m * 0.5 ) ), footer: tw( Math.max( 12, m * 0.4 ) ) }
            } },
            children
        };
        if ( h.left || h.right || h.logo ) {
            const kids = [ bar( h.left, h.right, { left: { bold: true, size: fs * 2 }, right: { color: hex6( MUTED ), size: ( fs - 1 ) * 2 } }, h.logo ) ];
            if ( h.line ) { kids.push( new D.Paragraph( { border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: hex6( LINE ), space: 1 } } } ) ); }
            section.headers = { default: new D.Header( { children: kids } ) };
        }
        if ( ft.left || ft.right ) {
            const o = { color: hex6( MUTED ), size: ( fs - 2 ) * 2 };
            section.footers = { default: new D.Footer( { children: [ bar( ft.left, ft.right, { left: o, right: o } ) ] } ) };
        }
        const document = new D.Document( {
            creator: "Report Designer by TechRevX",
            title: doc.name || "Report",
            styles: { default: { document: { run: { font: "Calibri", size: fs * 2, color: c.text }, paragraph: { spacing: { after: 0, line: 264 } } } } },
            sections: [ section ]
        } );
        return D.Packer.toBlob( document );
    }

    // image size in pixels for docx and Excel, from its width in points (or a box to fit)
    function imgBox( src, maxW, maxH, t, b ) {
        const ratio = b && b.ratio ? +b.ratio : ( t.header.logoRatio || 0.4 );
        let w = maxW, hgt = maxW * ratio;
        if ( maxH && hgt > maxH ) { hgt = maxH; w = maxH / ratio; }
        return { width: Math.round( w * PX ), height: Math.round( hgt * PX ) };
    }

    function docxNode( el, c ) {
        const D = c.D, b = el.b, tw = c.tw;
        const align = a => a === "center" ? D.AlignmentType.CENTER : a === "right" ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT;
        const runs = ( s, o ) => String( s ).split( "\n" ).map( ( line, i ) => new D.TextRun( Object.assign( { text: line, break: i ? 1 : 0 }, o ) ) );
        const none = { style: D.BorderStyle.NONE, size: 0, color: "FFFFFF" };
        const line = { style: D.BorderStyle.SINGLE, size: 4, color: hex6( LINE ) };
        switch ( el.type ) {
            case "heading": {
                const lvl = +b.level || 2;
                return new D.Paragraph( { alignment: align( b.align ), keepNext: true, spacing: { before: lvl === 1 ? 0 : tw( 8 ), after: tw( lvl === 1 ? 8 : 5 ) },
                    children: runs( el.text, { bold: true, size: ( HEADING_SIZES[ lvl ] || 15 ) * 2, color: b.color ? hex6( b.color ) : lvl === 1 ? c.text : c.accent } ) } );
            }
            case "text":
                return new D.Paragraph( { alignment: align( b.align ), spacing: { after: tw( 6 ) },
                    children: runs( el.text, { size: ( +b.size || c.fs ) * 2, bold: !!b.bold, italics: !!b.italic, color: b.color ? hex6( b.color ) : c.text } ) } );
            case "fields": {
                const n = Math.max( 1, Math.min( 3, +b.columns || 2 ) ), rows = [], right = b.align === "right";
                const total = right ? c.avail * Math.min( 0.5, 0.25 * n + 0.2 ) : c.avail, cellW = total / n;
                const nb = { top: none, bottom: none, left: none, right: none };
                for ( let i = 0; i < el.items.length; i += n ) {
                    const cells = [];
                    for ( let k = 0; k < n; k++ ) {
                        const it = el.items[ i + k ];
                        cells.push( new D.TableCell( { borders: nb, width: { size: tw( cellW * 0.4 ), type: D.WidthType.DXA },
                            children: [ new D.Paragraph( { children: runs( it ? it.label : "", { color: hex6( MUTED ), size: Math.round( ( c.fs - 0.5 ) * 2 ) } ) } ) ] } ) );
                        cells.push( new D.TableCell( { borders: nb, width: { size: tw( cellW * 0.6 ), type: D.WidthType.DXA },
                            children: [ new D.Paragraph( { alignment: right ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT, children: runs( it ? it.value : "", { bold: true } ) } ) ] } ) );
                    }
                    rows.push( new D.TableRow( { children: cells } ) );
                }
                if ( !rows.length ) { return null; }
                const widths = [];
                for ( let k = 0; k < n; k++ ) { widths.push( tw( cellW * 0.4 ), tw( cellW * 0.6 ) ); }
                return [ new D.Table( { rows, columnWidths: widths, width: { size: tw( total ), type: D.WidthType.DXA },
                                        alignment: right ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT,
                                        borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none } } ),
                         new D.Paragraph( { spacing: { after: tw( 6 ) } } ) ];
            }
            case "table": {
                if ( !el.cols.length ) { return null; }
                const n = el.cols.length, size = Math.round( ( +b.size || c.fs - 0.5 ) * 2 ), grid = b.border === "grid", noLines = b.border === "none";
                const widths = widthsOf( el.head.map( x => x.width ), c.avail ).map( tw );
                const tint = mix( c.accent, "FFFFFF", 0.88 );
                const margins = { top: 50, bottom: 50, left: 90, right: 90 };
                const cell = ( text, o, i, extra ) => new D.TableCell( Object.assign( {
                    width: { size: widths[ i ], type: D.WidthType.DXA }, margins,
                    shading: o.fill ? { type: D.ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
                    children: [ new D.Paragraph( { alignment: align( o.align ), children: runs( text, { bold: !!o.bold, size, color: o.color } ) } ) ]
                }, extra || {} ) );
                const rows = [ new D.TableRow( { tableHeader: true, cantSplit: true,
                    children: el.head.map( ( x, i ) => cell( x.text, { bold: true, color: "FFFFFF", fill: c.accent, align: x.align }, i ) ) } ) ];
                let zebra = 0;
                el.rows.forEach( r => {
                    if ( r.kind === "group" ) {
                        zebra = 0;
                        rows.push( new D.TableRow( { cantSplit: true, children: [ cell( r.text, { bold: true, color: c.accent, fill: tint }, 0,
                            { columnSpan: n, width: { size: widths.reduce( ( a, x ) => a + x, 0 ), type: D.WidthType.DXA } } ) ] } ) );
                    } else if ( r.kind === "subtotal" ) {
                        rows.push( new D.TableRow( { cantSplit: true, children: r.cells.map( ( x, i ) => cell( x.text, { bold: true, align: x.align }, i ) ) } ) );
                    } else {
                        const fill = b.striped !== false && zebra++ % 2 === 1 ? hex6( ZEBRA ) : null;
                        rows.push( new D.TableRow( { cantSplit: true, children: r.cells.map( ( x, i ) => cell( x.text, { align: x.align, fill }, i ) ) } ) );
                    }
                } );
                if ( el.total ) {
                    rows.push( new D.TableRow( { cantSplit: true, children: el.total.cells.map( ( x, i ) => cell( x.text, { bold: true, align: x.align, fill: tint }, i ) ) } ) );
                }
                return [ new D.Table( {
                    rows, columnWidths: widths, width: { size: widths.reduce( ( a, x ) => a + x, 0 ), type: D.WidthType.DXA }, layout: D.TableLayoutType.FIXED,
                    borders: { top: none, bottom: noLines ? none : line, left: grid ? line : none, right: grid ? line : none,
                               insideHorizontal: noLines ? none : line, insideVertical: grid ? line : none }
                } ), new D.Paragraph( { spacing: { after: tw( 8 ) } } ) ];
            }
            case "image": {
                if ( !b.src ) { return null; }
                const w = Math.min( +b.width || 140, c.avail );
                return new D.Paragraph( { alignment: align( b.align ), spacing: { after: tw( 8 ) }, children: [
                    new D.ImageRun( { type: imageType( b.src ), data: imageBytes( b.src ), transformation: { width: Math.round( w * PX ), height: Math.round( w * ( +b.ratio || 0.5 ) * PX ) } } ) ] } );
            }
            case "divider":
                return new D.Paragraph( { spacing: { before: tw( 4 ), after: tw( 10 ) },
                    border: { bottom: { style: D.BorderStyle.SINGLE, size: Math.max( 2, Math.round( ( +b.thickness || 1 ) * 8 ) ), color: b.color ? hex6( b.color ) : hex6( LINE ), space: 1 } } } );
            case "spacer":
                return new D.Paragraph( { spacing: { before: 0, after: 0, line: Math.max( 20, tw( +b.height || 16 ) ), lineRule: "exact" } } );
            case "pagebreak":
                return new D.Paragraph( { children: [ new D.PageBreak() ] } );
        }
        return null;
    }

    // ---------------------------------------------------------------------------------------------------------------
    // Excel (ExcelJS): one sheet, or one sheet per section; real numbers and dates, SUBTOTAL formulas for totals
    // ---------------------------------------------------------------------------------------------------------------
    function excelDateFormat( t, time ) {
        const f = ( t.settings.dateFormat || "DD MMM YYYY" ).replace( /YYYY|MMMM|MMM|MM|DD|D/g, x => ( { YYYY: "yyyy", MMMM: "mmmm", MMM: "mmm", MM: "mm", DD: "dd", D: "d" } )[ x ] );
        return time ? f + " hh:mm" : f;
    }
    function excelNumFmt( t, format, type ) {
        switch ( format ) {
            case "number": return "#,##0.00";
            case "integer": return "#,##0";
            case "currency": return "\"" + String( t.settings.currency || "" ).replace( /"/g, "" ) + "\"#,##0.00";
            case "percent": return "0.0%";
            case "date": return excelDateFormat( t );
            case "datetime": return excelDateFormat( t, true );
            case "text": return "@";
        }
        return type === "date" ? excelDateFormat( t ) : undefined;
    }
    // ExcelJS writes dates as UTC: give it the wall-clock time, so 2 Oct stays 2 Oct in every time zone
    const excelDate = d => new Date( Date.UTC( d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds() ) );
    const colLetter = n => { let s = ""; while ( n > 0 ) { const r = ( n - 1 ) % 26; s = String.fromCharCode( 65 + r ) + s; n = Math.floor( ( n - 1 ) / 26 ); } return s; };
    const SUBTOTAL_FN = { avg: 1, count: 3, max: 4, min: 5, sum: 9 };

    async function toXlsx( model, ExcelJS ) {
        const t = model.tpl, wb = new ExcelJS.Workbook(), m = +t.page.margin || 40, fs = +t.style.fontSize || 10;
        const accent = "FF" + hex6( t.style.accent, "0F766E" ), text = "FF" + hex6( t.style.text ), muted = "FF" + hex6( MUTED );
        wb.creator = "Report Designer by TechRevX";
        wb.created = new Date();
        const doc = model.doc, used = {};
        const hf = s => pageTokens( resolveText( s, doc ).replace( /&/g, "&&" ), "&P", "&N" );
        const sheetName = name => {
            let s = String( name || "Report" ).replace( /[\\/?*[\]:]/g, " " ).trim().slice( 0, 28 ) || "Report", k = s, i = 2;
            while ( used[ k.toLowerCase() ] ) { k = s.slice( 0, 26 ) + " " + i++; }
            used[ k.toLowerCase() ] = true;
            return k;
        };
        const tables = model.groups[ 0 ].els.filter( e => e.type === "table" );
        const width = Math.max( 4, ...tables.map( e => e.cols.length ), ...model.groups[ 0 ].els.filter( e => e.type === "fields" ).map( e => 2 * ( +e.b.columns || 2 ) ) );
        const newSheet = name => {
            const ws = wb.addWorksheet( sheetName( name ), {
                views: [ { showGridLines: false } ],
                pageSetup: { paperSize: EXCEL_PAPER[ t.page.size ] || 9, orientation: t.page.orientation === "landscape" ? "landscape" : "portrait",
                             fitToPage: true, fitToWidth: 1, fitToHeight: 0,
                             margins: { left: m / 72, right: m / 72, top: ( m + 18 ) / 72, bottom: ( m + 14 ) / 72, header: m / 144, footer: m / 144 } },
                headerFooter: {
                    oddHeader: t.header.left || t.header.right ? "&L&B" + hf( t.header.left ) + "&B&R" + hf( t.header.right ) : undefined,
                    oddFooter: t.footer.left || t.footer.right ? "&L" + hf( t.footer.left ) + "&R" + hf( t.footer.right ) : undefined
                }
            } );
            // column widths from the first table, in characters
            const first = tables[ 0 ];
            for ( let i = 1; i <= width; i++ ) {
                ws.getColumn( i ).width = first && first.head[ i - 1 ] ? Math.round( 13 * first.head[ i - 1 ].width + 3 ) : 16;
            }
            ws.properties.defaultRowHeight = Math.round( fs * 1.5 );
            return { ws, r: 1, filter: false };
        };
        const perSheet = model.groupBy && t.settings.sheetPerGroup !== false && model.groups.length > 1;
        let sh = newSheet( perSheet ? model.f.value( model.groups[ 0 ].key, model.doc.data.types[ model.groupBy ], "" ) : doc.name );
        model.groups.forEach( ( g, gi ) => {
            if ( gi > 0 ) {
                if ( perSheet ) {
                    sh = newSheet( model.f.value( g.key, model.doc.data.types[ model.groupBy ], "" ) );
                } else {
                    if ( t.settings.groupPageBreak ) { sh.ws.getRow( sh.r - 1 ).addPageBreak(); }
                    sh.r += 1;
                }
            }
            g.els.forEach( el => xlsxNode( el, sh, { t, wb, width, fs, accent, text, muted } ) );
        } );
        return new Blob( [ await wb.xlsx.writeBuffer() ], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } );
    }

    function xlsxNode( el, sh, c ) {
        const ws = sh.ws, b = el.b, W = c.width;
        const font = o => Object.assign( { name: "Calibri", size: c.fs, color: { argb: c.text } }, o );
        const horiz = a => a === "center" ? "center" : a === "right" ? "right" : "left";
        const merged = ( value, f, align, lines ) => {
            ws.mergeCells( sh.r, 1, sh.r, W );
            const cell = ws.getCell( sh.r, 1 );
            cell.value = value;
            cell.font = f;
            cell.alignment = { horizontal: horiz( align ), vertical: "top", wrapText: true };
            ws.getRow( sh.r ).height = Math.max( 15, ( f.size || c.fs ) * 1.45 * ( lines || 1 ) );
            sh.r++;
        };
        const thin = { style: "thin", color: { argb: "FF" + hex6( LINE ) } };
        switch ( el.type ) {
            case "heading": {
                const lvl = +b.level || 2;
                merged( el.text, font( { size: HEADING_SIZES[ lvl ] || 15, bold: true, color: { argb: b.color ? "FF" + hex6( b.color ) : lvl === 1 ? c.text : c.accent } } ), b.align );
                return;
            }
            case "text": {
                const size = +b.size || c.fs, chars = W * 14 * 10 / size;
                const lines = String( el.text ).split( "\n" ).reduce( ( a, l ) => a + Math.max( 1, Math.ceil( l.length / chars ) ), 0 );
                merged( el.text, font( { size, bold: !!b.bold, italic: !!b.italic, color: { argb: b.color ? "FF" + hex6( b.color ) : c.text } } ), b.align, lines );
                return;
            }
            case "fields": {
                const n = Math.max( 1, Math.min( 3, +b.columns || 2 ) ), start = b.align === "right" ? Math.max( 1, W - 2 * n + 1 ) : 1;
                el.items.forEach( ( it, i ) => {
                    const row = sh.r + Math.floor( i / n ), col = start + ( i % n ) * 2;
                    const l = ws.getCell( row, col ), v = ws.getCell( row, col + 1 );
                    l.value = it.label;
                    l.font = font( { color: { argb: c.muted } } );
                    v.value = it.value;
                    v.font = font( { bold: true } );
                    v.alignment = { horizontal: b.align === "right" ? "right" : "left" };
                } );
                sh.r += Math.ceil( el.items.length / n ) + 1;
                return;
            }
            case "table": {
                if ( !el.cols.length ) { return; }
                const t = c.t, n = el.cols.length, tint = "FF" + mix( c.accent.slice( 2 ), "FFFFFF", 0.88 ), grid = b.border === "grid", noLines = b.border === "none";
                const headRow = sh.r;
                el.head.forEach( ( h, i ) => {
                    const cell = ws.getCell( sh.r, i + 1 );
                    cell.value = h.text;
                    cell.font = font( { bold: true, color: { argb: "FFFFFFFF" } } );
                    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.accent } };
                    cell.alignment = { horizontal: horiz( h.align ), vertical: "middle" };
                } );
                ws.getRow( sh.r ).height = Math.round( c.fs * 1.9 );
                sh.r++;
                const firstData = sh.r;
                let zebra = 0, groupStart = sh.r, groups = false;
                const border = ( cell, top ) => {
                    if ( noLines ) { return; }
                    cell.border = grid ? { top: thin, left: thin, bottom: thin, right: thin } : { bottom: thin, top: top ? thin : undefined };
                };
                const writeTotals = ( r, from, to, fill ) => {
                    r.cells.forEach( ( x, i ) => {
                        const cell = ws.getCell( sh.r, i + 1 );
                        if ( x.total && SUBTOTAL_FN[ x.total ] && to >= from ) {
                            const range = colLetter( i + 1 ) + from + ":" + colLetter( i + 1 ) + to;
                            cell.value = { formula: "SUBTOTAL(" + SUBTOTAL_FN[ x.total ] + "," + range + ")",
                                           result: typeof x.v === "number" ? ( x.format === "percent" ? x.v / 100 : x.v ) : x.v instanceof Date ? excelDate( x.v ) : undefined };
                            cell.numFmt = excelNumFmt( t, x.format, x.type ) || ( x.total === "avg" ? "#,##0.00" : undefined );
                        } else {
                            cell.value = x.text;
                        }
                        cell.font = font( { bold: true } );
                        cell.alignment = { horizontal: horiz( x.align ) };
                        if ( fill ) { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } }; }
                        border( cell, true );
                    } );
                    sh.r++;
                };
                el.rows.forEach( r => {
                    if ( r.kind === "group" ) {
                        groups = true;
                        zebra = 0;
                        ws.mergeCells( sh.r, 1, sh.r, n );
                        const cell = ws.getCell( sh.r, 1 );
                        cell.value = r.text;
                        cell.font = font( { bold: true, color: { argb: c.accent } } );
                        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: tint } };
                        sh.r++;
                        groupStart = sh.r;
                    } else if ( r.kind === "subtotal" ) {
                        writeTotals( r, groupStart, sh.r - 1 );
                    } else {
                        const fill = b.striped !== false && zebra++ % 2 === 1 ? "FF" + hex6( ZEBRA ) : null;
                        r.cells.forEach( ( x, i ) => {
                            const cell = ws.getCell( sh.r, i + 1 );
                            let v = x.v;
                            if ( v === null || v === undefined ) { v = null; }
                            else if ( x.format === "text" ) { v = x.text; }
                            else if ( typeof v === "number" && x.format === "percent" ) { v = v / 100; }
                            else if ( x.type === "number" && typeof v !== "number" && !isNaN( Number( v ) ) && v !== "" ) { v = Number( v ); }
                            if ( v instanceof Date ) { v = excelDate( v ); }
                            cell.value = v;
                            const nf = excelNumFmt( t, x.format, x.type );
                            if ( nf && !( v instanceof Date && x.format === "" ) ) { cell.numFmt = nf; } else if ( v instanceof Date ) { cell.numFmt = excelDateFormat( t ); }
                            cell.alignment = { horizontal: horiz( x.align ), vertical: "top" };
                            cell.font = font();
                            if ( fill ) { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } }; }
                            border( cell );
                        } );
                        sh.r++;
                    }
                } );
                if ( el.total ) { writeTotals( el.total, firstData, sh.r - 1, tint ); }
                if ( !groups && !sh.filter && sh.r - 1 > firstData ) {
                    ws.autoFilter = { from: { row: headRow, column: 1 }, to: { row: firstData + el.rows.length - 1, column: n } };
                    sh.filter = true;
                }
                sh.r++;
                return;
            }
            case "image": {
                if ( !b.src ) { return; }
                const w = ( +b.width || 140 ) * PX, h = w * ( +b.ratio || 0.5 );
                const id = c.wb.addImage( { base64: b.src, extension: imageType( b.src ) === "png" ? "png" : "jpeg" } );
                const colPx = 7 * ( ws.getColumn( 1 ).width || 16 ) + 5;
                const col = b.align === "right" ? Math.max( 0, W - w / colPx ) : b.align === "center" ? Math.max( 0, ( W - w / colPx ) / 2 ) : 0;
                ws.addImage( id, { tl: { col, row: sh.r - 1 }, ext: { width: Math.round( w ), height: Math.round( h ) }, editAs: "oneCell" } );
                sh.r += Math.ceil( h / 20 ) + 1;
                return;
            }
            case "divider":
                for ( let i = 1; i <= W; i++ ) {
                    ws.getCell( sh.r, i ).border = { bottom: { style: ( +b.thickness || 1 ) > 1.5 ? "medium" : "thin", color: { argb: "FF" + hex6( b.color || LINE ) } } };
                }
                ws.getRow( sh.r ).height = 8;
                sh.r += 2;
                return;
            case "spacer":
                ws.getRow( sh.r ).height = Math.max( 4, +b.height || 16 );
                sh.r++;
                return;
            case "pagebreak":
                if ( sh.r > 1 ) { ws.getRow( sh.r - 1 ).addPageBreak(); }
                return;
        }
    }

    // ---------------------------------------------------------------------------------------------------------------
    // preview (HTML)
    // ---------------------------------------------------------------------------------------------------------------
    function blockHtml( el, ctx, names ) {
        const b = el.b, rich = s => resolveHtml( s, ctx, names );
        const style = o => Object.keys( o ).filter( k => o[ k ] !== "" && o[ k ] !== undefined && o[ k ] !== null ).map( k => k + ":" + o[ k ] ).join( ";" );
        switch ( el.type ) {
            case "heading": {
                const lvl = +b.level || 2;
                return "<div class=\"tx-Rd-h tx-Rd-h" + lvl + "\" style=\"" + style( { "text-align": b.align, color: b.color ? "#" + hex6( b.color ) : "" } ) + "\">" + ( rich( b.text ) || "&nbsp;" ) + "</div>";
            }
            case "text":
                return "<div class=\"tx-Rd-text\" style=\"" + style( { "text-align": b.align, color: b.color ? "#" + hex6( b.color ) : "", "font-size": +b.size ? b.size + "pt" : "",
                    "font-weight": b.bold ? "700" : "", "font-style": b.italic ? "italic" : "" } ) + "\">" + ( rich( b.text ) || "&nbsp;" ) + "</div>";
            case "fields": {
                const n = Math.max( 1, Math.min( 3, +b.columns || 2 ) ), items = b.items || [];
                if ( !items.length ) { return "<div class=\"tx-Rd-empty\">Add fields in the panel on the right.</div>"; }
                let rows = "";
                for ( let i = 0; i < items.length; i += n ) {
                    rows += "<tr>";
                    for ( let k = 0; k < n; k++ ) {
                        const it = items[ i + k ];
                        rows += "<th scope=\"row\">" + ( it ? rich( it.label ) : "" ) + "</th><td>" + ( it ? rich( it.value ) : "" ) + "</td>";
                    }
                    rows += "</tr>";
                }
                return "<table class=\"tx-Rd-fields" + ( b.align === "right" ? " is-right" : "" ) + "\">" + rows + "</table>";
            }
            case "table": {
                if ( !el.cols.length ) { return "<div class=\"tx-Rd-empty\">Choose the table's columns in the panel on the right.</div>"; }
                const sum = el.head.reduce( ( a, h ) => a + h.width, 0 ) || 1;
                const cols = "<colgroup>" + el.head.map( h => "<col style=\"width:" + ( 100 * h.width / sum ).toFixed( 2 ) + "%\">" ).join( "" ) + "</colgroup>";
                const head = "<thead><tr>" + el.cols.map( ( c, i ) => "<th style=\"text-align:" + ( c.align || "left" ) + "\">" + rich( c.label ) + "</th>" ).join( "" ) + "</tr></thead>";
                const cellsHtml = ( cells, tag ) => cells.map( x => "<" + tag + " style=\"text-align:" + x.align + "\">" + esc( x.text ) + "</" + tag + ">" ).join( "" );
                let zebra = 0, body = "";
                el.rows.forEach( r => {
                    if ( r.kind === "group" ) { zebra = 0; body += "<tr class=\"is-group\"><td colspan=\"" + el.cols.length + "\">" + esc( r.text ) + "</td></tr>"; }
                    else if ( r.kind === "subtotal" ) { body += "<tr class=\"is-subtotal\">" + cellsHtml( r.cells, "td" ) + "</tr>"; }
                    else { body += "<tr" + ( zebra++ % 2 === 1 && b.striped !== false ? " class=\"is-zebra\"" : "" ) + ">" + cellsHtml( r.cells, "td" ) + "</tr>"; }
                } );
                if ( !el.rows.length ) { body = "<tr><td class=\"tx-Rd-noRows\" colspan=\"" + el.cols.length + "\">No rows</td></tr>"; }
                const foot = el.total ? "<tfoot><tr>" + cellsHtml( el.total.cells, "td" ) + "</tr></tfoot>" : "";
                return "<table class=\"tx-Rd-table is-" + ( b.border || "lines" ) + "\"" + ( +b.size ? " style=\"font-size:" + b.size + "pt\"" : "" ) + ">" + cols + head + "<tbody>" + body + "</tbody>" + foot + "</table>"
                    + ( el.hidden ? "<div class=\"tx-Rd-moreRows\">" + esc( el.hidden ) + " more rows in the downloaded report</div>" : "" );
            }
            case "image":
                return b.src
                    ? "<div style=\"text-align:" + ( b.align || "left" ) + "\"><img class=\"tx-Rd-img\" alt=\"\" src=\"" + esc( b.src ) + "\" style=\"width:" + ( +b.width || 140 ) + "pt\"></div>"
                    : "<div class=\"tx-Rd-empty\"><span class=\"fa fa-image\" aria-hidden=\"true\"></span> Choose an image in the panel on the right.</div>";
            case "divider":
                return "<hr class=\"tx-Rd-hr\" style=\"border-top:" + ( +b.thickness || 1 ) + "pt solid " + ( b.color ? "#" + hex6( b.color ) : LINE ) + "\">";
            case "spacer":
                return "<div class=\"tx-Rd-spacer\" style=\"height:" + ( +b.height || 16 ) + "pt\"></div>";
            case "pagebreak":
                return "<div class=\"tx-Rd-break\"><span>Page break</span></div>";
        }
        return "";
    }

    // ---------------------------------------------------------------------------------------------------------------
    // the designer
    // ---------------------------------------------------------------------------------------------------------------
    const ICON = ( name, label ) => "<span class=\"fa " + name + "\" aria-hidden=\"true\"></span>" + ( label ? "<span class=\"tx-Rd-btnText\">" + label + "</span>" : "" );
    const FORMAT_INFO = {
        pdf: { label: "PDF", icon: "fa-file-pdf-o", ext: "pdf" },
        docx: { label: "Word", icon: "fa-file-word-o", ext: "docx" },
        xlsx: { label: "Excel", icon: "fa-file-excel-o", ext: "xlsx" }
    };
    const options = ( list, value ) => list.map( o => "<option value=\"" + esc( o[ 0 ] ) + "\"" + ( String( o[ 0 ] ) === String( value === undefined || value === null ? "" : value ) ? " selected" : "" ) + ">" + esc( o[ 1 ] ) + "</option>" ).join( "" );

    function getPath( obj, path ) { return path.split( "." ).reduce( ( o, k ) => o === undefined || o === null ? undefined : o[ k ], obj ); }
    function setPath( obj, path, value ) {
        const keys = path.split( "." ), last = keys.pop();
        const target = keys.reduce( ( o, k ) => o[ k ], obj );
        target[ last ] = value;
    }

    class Designer {
        constructor( regionId ) {
            this.id = regionId;
            this.el = document.getElementById( regionId + "_report" );
            this.cfg = JSON.parse( document.getElementById( regionId + "_report_cfg" ).textContent );
            this.editable = this.cfg.mode !== "report";
            this.templates = this.cfg.templates || [];
            this.formats = String( this.cfg.formats || "pdf:docx:xlsx" ).split( ":" ).filter( f => FORMAT_INFO[ f ] );
            this.data = new Data( null );
            this.meta = { id: null, name: "", description: "", shared: false, mine: true };
            this.tpl = normalize( null );
            this.selected = null;
            this.names = false;
            this.undoStack = [];
            this.redoStack = [];
            this.tab = "report";
            this.build();
            this.snapshot( true );
            this.registerRegion();
            this.loadData().then( () => this.openInitial() );
        }

        // ---------- markup ----------
        build() {
            const id = this.id, ed = this.editable;
            this.el.classList.toggle( "tx-Rd--report", !ed );
            this.el.innerHTML =
                "<div class=\"tx-Rd-bar\" role=\"toolbar\" aria-label=\"" + ( ed ? "Report designer" : "Reports" ) + "\">"
                + "<div class=\"tx-Rd-group tx-Rd-tplGroup\">"
                +   "<label class=\"tx-Rd-srOnly\" for=\"" + id + "_tpl\">Template</label>"
                +   "<span class=\"tx-Rd-selectWrap\"><select id=\"" + id + "_tpl\" class=\"tx-Rd-select tx-Rd-tplSelect\"></select></span>"
                +   ( ed ? "<button type=\"button\" class=\"tx-Rd-btn\" data-act=\"new\" title=\"New template\">" + ICON( "fa-plus", "New" ) + "</button>"
                         + "<button type=\"button\" class=\"tx-Rd-btn tx-Rd-btn--primary\" data-act=\"save\" title=\"Save (Ctrl+S)\">" + ICON( "fa-save", "Save" ) + "</button>"
                         + "<button type=\"button\" class=\"tx-Rd-btn\" data-act=\"saveas\" title=\"Save as a new template\">" + ICON( "fa-copy", "Save as" ) + "</button>"
                         + "<button type=\"button\" class=\"tx-Rd-btn tx-Rd-btn--icon\" data-act=\"delete\" title=\"Delete template\" aria-label=\"Delete template\">" + ICON( "fa-trash-o" ) + "</button>"
                         + "<span class=\"tx-Rd-status\" role=\"status\" aria-live=\"polite\"></span>" : "" )
                + "</div>"
                + ( ed ? "<div class=\"tx-Rd-group\">"
                         + "<button type=\"button\" class=\"tx-Rd-btn tx-Rd-btn--icon\" data-act=\"undo\" title=\"Undo (Ctrl+Z)\" aria-label=\"Undo\">" + ICON( "fa-undo" ) + "</button>"
                         + "<button type=\"button\" class=\"tx-Rd-btn tx-Rd-btn--icon\" data-act=\"redo\" title=\"Redo (Ctrl+Y)\" aria-label=\"Redo\">" + ICON( "fa-repeat" ) + "</button>"
                         + "<div class=\"tx-Rd-seg\" role=\"group\" aria-label=\"Show in the preview\">"
                         +   "<button type=\"button\" data-act=\"view\" data-view=\"data\" aria-pressed=\"true\" title=\"Show the values of the first rows\">Data</button>"
                         +   "<button type=\"button\" data-act=\"view\" data-view=\"names\" aria-pressed=\"false\" title=\"Show the field names\">Fields</button>"
                         + "</div></div>" : "" )
                + "<div class=\"tx-Rd-group tx-Rd-downloads\"><span class=\"tx-Rd-dlLabel\">Download</span>"
                +   this.formats.map( f => "<button type=\"button\" class=\"tx-Rd-btn tx-Rd-dl tx-Rd-dl--" + f + "\" data-act=\"download\" data-format=\"" + f + "\" title=\"Download as " + FORMAT_INFO[ f ].label + "\">"
                                          + ICON( FORMAT_INFO[ f ].icon, FORMAT_INFO[ f ].label ) + "</button>" ).join( "" )
                + "</div></div>"
                + "<div class=\"tx-Rd-body\">"
                + ( ed ? "<aside class=\"tx-Rd-side tx-Rd-left\" aria-label=\"Blocks and fields\">"
                         + "<h3 class=\"tx-Rd-sideTitle\">Add a block</h3>"
                         + "<p class=\"tx-Rd-sideHint\">Drag onto the page, or click to add below the selected block.</p>"
                         // tiles are not <button>s: Firefox does not drag buttons
                         + "<div class=\"tx-Rd-palette\">" + Object.keys( BLOCKS ).map( k => "<div role=\"button\" tabindex=\"0\" class=\"tx-Rd-tile\" draggable=\"true\" data-add=\"" + k + "\" title=\"" + esc( BLOCKS[ k ].hint ) + "\">"
                             + "<span class=\"fa " + BLOCKS[ k ].icon + "\" aria-hidden=\"true\"></span><span>" + BLOCKS[ k ].label + "</span></div>" ).join( "" ) + "</div>"
                         + "<h3 class=\"tx-Rd-sideTitle\">Fields</h3>"
                         + "<p class=\"tx-Rd-sideHint\">Click to insert into the selected text box, or drag into it.</p>"
                         + "<input type=\"search\" class=\"tx-Rd-input tx-Rd-fieldSearch\" placeholder=\"Find a field\" aria-label=\"Find a field\">"
                         + "<ul class=\"tx-Rd-fieldList\"></ul>"
                         + "</aside>" : "" )
                + "<div class=\"tx-Rd-stage\"><div class=\"tx-Rd-pages\"></div><div class=\"tx-Rd-drop\" hidden></div></div>"
                + ( ed ? "<aside class=\"tx-Rd-side tx-Rd-right\" aria-label=\"Properties\">"
                         + "<div class=\"tx-Rd-tabs\" role=\"tablist\">"
                         +   "<button type=\"button\" role=\"tab\" id=\"" + id + "_tabBlock\" aria-controls=\"" + id + "_props\" data-tab=\"block\" aria-selected=\"false\">Block</button>"
                         +   "<button type=\"button\" role=\"tab\" id=\"" + id + "_tabReport\" aria-controls=\"" + id + "_props\" data-tab=\"report\" aria-selected=\"true\">Report</button>"
                         + "</div><div class=\"tx-Rd-props\" id=\"" + id + "_props\" role=\"tabpanel\"></div></aside>" : "" )
                + "</div>"
                + "<dialog class=\"tx-Rd-dialog\" aria-labelledby=\"" + id + "_dlgTitle\"></dialog>"
                + "<div class=\"tx-Rd-pop\" hidden></div>";

            const q = s => this.el.querySelector( s );
            this.bar = q( ".tx-Rd-bar" );
            this.tplSelect = q( ".tx-Rd-tplSelect" );
            this.stage = q( ".tx-Rd-stage" );
            this.pagesEl = q( ".tx-Rd-pages" );
            this.dropEl = q( ".tx-Rd-drop" );
            this.props = q( ".tx-Rd-props" );
            this.fieldList = q( ".tx-Rd-fieldList" );
            this.dialog = q( ".tx-Rd-dialog" );
            // the page is one big <form>, and a form inside it is dropped by the browser: the dialog lives in <body>
            document.body.appendChild( this.dialog );
            this.pop = q( ".tx-Rd-pop" );
            this.status = q( ".tx-Rd-status" );
            this.bind();
            this.fillTemplateSelect();
            new ResizeObserver( () => this.fit() ).observe( this.stage );
        }

        bind() {
            this.bar.addEventListener( "click", e => {
                const btn = e.target.closest( "button[data-act]" );
                if ( !btn || btn.disabled ) { return; }
                switch ( btn.dataset.act ) {
                    case "new": this.newDialog(); break;
                    case "save": this.save(); break;
                    case "saveas": this.saveDialog( true ); break;
                    case "delete": this.remove(); break;
                    case "undo": this.undo(); break;
                    case "redo": this.redo(); break;
                    case "view": this.names = btn.dataset.view === "names"; this.updateBar(); this.render(); break;
                    case "download": this.download( btn.dataset.format, btn ); break;
                }
            } );
            this.tplSelect.addEventListener( "change", () => {
                const v = this.tplSelect.value;
                this.confirmDiscard( () => this.openTemplate( v ), () => { this.tplSelect.value = this.meta.id || ""; } );
            } );
            if ( !this.editable ) { return; }

            // palette: click adds below the selection, drag drops where the line shows
            const palette = this.el.querySelector( ".tx-Rd-palette" );
            palette.addEventListener( "click", e => { const t = e.target.closest( "[data-add]" ); if ( t ) { this.addBlock( t.dataset.add ); } } );
            palette.addEventListener( "dragstart", e => {
                const t = e.target.closest( "[data-add]" );
                if ( !t ) { return; }
                e.dataTransfer.setData( "application/x-tx-block", t.dataset.add );
                e.dataTransfer.effectAllowed = "copy";
                this.dragging = { add: t.dataset.add };
            } );
            // fields: click inserts into the last focused text box; drag drops "{FIELD}" into any text box
            this.fieldList.addEventListener( "click", e => { const li = e.target.closest( "[data-field]" ); if ( li ) { this.insertField( "{" + li.dataset.field + "}" ); } } );
            // the tiles and fields are role="button": Enter and Space press them
            this.el.querySelector( ".tx-Rd-left" ).addEventListener( "keydown", e => {
                if ( ( e.key === "Enter" || e.key === " " ) && e.target.matches( "[role=button]" ) ) { e.preventDefault(); e.target.click(); }
            } );
            this.fieldList.addEventListener( "dragstart", e => {
                const li = e.target.closest( "[data-field]" );
                if ( li ) { e.dataTransfer.setData( "text/plain", "{" + li.dataset.field + "}" ); e.dataTransfer.effectAllowed = "copy"; }
            } );
            this.el.querySelector( ".tx-Rd-fieldSearch" ).addEventListener( "input", e => this.renderFields( e.target.value ) );

            // the page: select, keyboard, block toolbar, drag to move
            this.stage.addEventListener( "click", e => {
                const act = e.target.closest( "[data-block-act]" );
                if ( act ) { this.blockAction( act.dataset.blockAct, act.closest( "[data-id]" ).dataset.id ); return; }
                const slot = e.target.closest( ".tx-Rd-slot" );
                if ( slot ) { this.select( slot.dataset.id ); } else if ( e.target.closest( ".tx-Rd-pHead, .tx-Rd-pFoot" ) ) { this.select( null ); this.setTab( "report", "header" ); } else if ( !e.target.closest( ".tx-Rd-blocks" ) ) { this.select( null ); }
            } );
            this.stage.addEventListener( "keydown", e => {
                const slot = e.target.closest( ".tx-Rd-slot" );
                if ( !slot || e.target.closest( ".tx-Rd-blockBar" ) ) { return; }
                const id = slot.dataset.id;
                if ( e.key === "Enter" || e.key === " " ) { e.preventDefault(); this.select( id ); this.focusProps(); }
                else if ( ( e.key === "Delete" || e.key === "Backspace" ) ) { e.preventDefault(); this.blockAction( "delete", id ); }
                else if ( e.altKey && ( e.key === "ArrowUp" || e.key === "ArrowDown" ) ) { e.preventDefault(); this.blockAction( e.key === "ArrowUp" ? "up" : "down", id ); }
                else if ( e.key === "ArrowUp" || e.key === "ArrowDown" ) {
                    e.preventDefault();
                    const i = this.index( id ) + ( e.key === "ArrowUp" ? -1 : 1 );
                    if ( this.tpl.blocks[ i ] ) { this.select( this.tpl.blocks[ i ].id ); this.focusBlock( this.tpl.blocks[ i ].id ); }
                }
                else if ( e.key === "Escape" ) { this.select( null ); }
            } );
            this.stage.addEventListener( "dragstart", e => {
                const h = e.target.closest( "[data-block-act=\"move\"]" );
                if ( !h ) { return; }
                const slot = h.closest( ".tx-Rd-slot" );
                e.dataTransfer.setData( "application/x-tx-move", slot.dataset.id );
                e.dataTransfer.effectAllowed = "move";
                try { e.dataTransfer.setDragImage( slot, 20, 20 ); } catch ( x ) { /* default image */ }
                this.dragging = { move: slot.dataset.id };
            } );
            this.stage.addEventListener( "dragover", e => {
                if ( !this.dragging ) { return; }
                e.preventDefault();
                e.dataTransfer.dropEffect = this.dragging.move ? "move" : "copy";
                this.showDrop( e.clientY, e.target );
            } );
            this.stage.addEventListener( "dragleave", e => { if ( !this.stage.contains( e.relatedTarget ) ) { this.hideDrop(); } } );
            this.stage.addEventListener( "drop", e => {
                if ( !this.dragging ) { return; }
                e.preventDefault();
                const at = this.dropIndex;
                this.hideDrop();
                if ( at === undefined || at === null ) { return; }
                if ( this.dragging.add ) { this.addBlock( this.dragging.add, at ); }
                else if ( this.dragging.move ) { this.moveBlock( this.dragging.move, at ); }
                this.dragging = null;
            } );
            document.addEventListener( "dragend", () => { this.dragging = null; this.hideDrop(); } );

            // properties
            this.el.querySelector( ".tx-Rd-tabs" ).addEventListener( "click", e => { const t = e.target.closest( "[data-tab]" ); if ( t ) { this.setTab( t.dataset.tab ); } } );
            this.el.querySelector( ".tx-Rd-tabs" ).addEventListener( "keydown", e => {
                if ( e.key === "ArrowLeft" || e.key === "ArrowRight" ) { this.setTab( this.tab === "block" ? "report" : "block" ); this.el.querySelector( "[data-tab=\"" + this.tab + "\"]" ).focus(); }
            } );
            this.props.addEventListener( "input", e => this.propInput( e, false ) );
            this.props.addEventListener( "change", e => this.propInput( e, true ) );
            this.props.addEventListener( "click", e => this.propClick( e ) );
            this.props.addEventListener( "focusin", e => { if ( e.target.matches( "[data-fields-ok]" ) ) { this.lastInput = e.target; } } );

            // shortcuts
            this.el.addEventListener( "keydown", e => {
                if ( !( e.ctrlKey || e.metaKey ) || e.altKey ) { return; }
                const k = e.key.toLowerCase();
                if ( k === "s" ) { e.preventDefault(); this.save(); }
                else if ( k === "z" && !e.shiftKey && !e.target.matches( "input, textarea" ) ) { e.preventDefault(); this.undo(); }
                else if ( ( k === "y" || ( k === "z" && e.shiftKey ) ) && !e.target.matches( "input, textarea" ) ) { e.preventDefault(); this.redo(); }
            } );
            document.addEventListener( "mousedown", e => { if ( !this.pop.hidden && !this.pop.contains( e.target ) && !e.target.closest( "[data-insert]" ) ) { this.closePop(); } } );
        }

        // ---------- data and templates ----------
        ajax( x01, extra, opts ) {
            const data = Object.assign( { x01 }, extra || {} );
            if ( this.cfg.itemsToSubmit ) { data.pageItems = this.cfg.itemsToSubmit; }
            return new Promise( ( resolve, reject ) => {
                apex.server.plugin( this.cfg.ajaxIdentifier, data, Object.assign( { dataType: "json" }, opts || {} ) )
                    .done( res => res && res.error ? reject( new Error( res.error ) ) : resolve( res ) )
                    .fail( ( xhr, status, err ) => {
                        const r = xhr && xhr.responseJSON;
                        reject( new Error( ( r && ( r.error || r.message ) ) || err || "The request failed." ) );
                    } );
            } );
        }

        showError( e ) {
            apex.message.clearErrors();
            apex.message.showErrors( [ { type: "error", location: "page", message: String( e && e.message || e ).replace( /^ORA-\d+:\s*/, "" ), unsafe: true } ] );
        }

        loadData() {
            const spin = apex.util.showSpinner( $( this.stage ) );
            return this.ajax( "DATA" ).then( res => {
                this.data = new Data( res );
                this.renderFields();
                this.render();
            } ).catch( e => this.showError( e ) ).finally( () => spin.remove() );
        }

        openInitial() {
            const want = this.cfg.templateId, list = this.templates;
            const pick = list.find( t => String( t.id ) === String( want ) ) || list.find( t => t.mine ) || list[ 0 ];
            if ( pick ) { return this.openTemplate( pick.id ); }
            if ( this.editable ) {
                this.setTemplate( starter( "list", this.data ), { id: null, name: "New report", description: "", shared: false, mine: true } );
                this.setStatus( "New template, not saved yet" );
            } else {
                this.render();
            }
        }

        fillTemplateSelect() {
            const mine = this.templates.filter( t => t.mine ), shared = this.templates.filter( t => !t.mine );
            const opt = t => "<option value=\"" + esc( t.id ) + "\">" + esc( t.name ) + ( t.shared && t.mine ? " (shared)" : "" ) + "</option>";
            let html = "";
            if ( !this.meta.id && this.editable ) { html += "<option value=\"\">" + esc( this.meta.name || "New report" ) + " (not saved)</option>"; }
            if ( !this.templates.length && !this.editable ) { html += "<option value=\"\">No templates yet</option>"; }
            if ( mine.length ) { html += shared.length ? "<optgroup label=\"My templates\">" + mine.map( opt ).join( "" ) + "</optgroup>" : mine.map( opt ).join( "" ); }
            if ( shared.length ) { html += "<optgroup label=\"Shared with me\">" + shared.map( t => "<option value=\"" + esc( t.id ) + "\">" + esc( t.name ) + "</option>" ).join( "" ) + "</optgroup>"; }
            this.tplSelect.innerHTML = html;
            this.tplSelect.value = this.meta.id ? String( this.meta.id ) : "";
            this.tplSelect.disabled = !this.templates.length && !this.editable;
        }

        openTemplate( id ) {
            if ( !id ) { return Promise.resolve(); }
            return this.ajax( "LOAD", { x02: id } ).then( res => {
                let tpl = null;
                try { tpl = JSON.parse( res.template || "{}" ); } catch ( e ) { apex.debug.warn( "Report Designer: the template is not valid JSON" ); }
                this.setTemplate( tpl, { id: res.id, name: res.name, description: res.description || "", shared: !!res.shared, mine: !!res.mine, owner: res.owner } );
                if ( this.cfg.templateItem ) { apex.item( this.cfg.templateItem ).setValue( String( res.id ) ); }
            } ).catch( e => this.showError( e ) );
        }

        setTemplate( tpl, meta ) {
            this.tpl = normalize( tpl );
            if ( meta ) { this.meta = meta; }
            this.selected = null;
            this.undoStack = [];
            this.redoStack = [];
            this.snapshot( true );
            this.fillTemplateSelect();
            this.setTab( "report" );
            this.render();
            this.updateBar();
            this.setStatus( this.meta.mine ? "" : "Shared, read only. Save as makes your copy." );
        }

        // ---------- history ----------
        state() { return JSON.stringify( { t: this.tpl, n: this.meta.name, d: this.meta.description, s: this.meta.shared } ); }
        snapshot( saved ) {
            this.current = this.state();
            if ( saved ) { this.saved = this.current; }
        }
        commit( typing ) {
            const now = this.state();
            if ( now === this.current ) { return; }
            if ( !( typing && this.lastTyping && Date.now() - this.lastTyping < 900 ) ) {
                this.undoStack.push( this.current );
                if ( this.undoStack.length > 100 ) { this.undoStack.shift(); }
            }
            this.lastTyping = typing ? Date.now() : 0;
            this.redoStack = [];
            this.current = now;
            this.updateBar();
        }
        restore( s ) {
            const o = JSON.parse( s );
            this.tpl = normalize( o.t );
            this.meta.name = o.n;
            this.meta.description = o.d;
            this.meta.shared = o.s;
            this.current = s;
            if ( this.selected && !this.block( this.selected ) ) { this.selected = null; }
            this.render();
            this.renderProps();
            this.updateBar();
        }
        undo() { if ( this.undoStack.length ) { this.redoStack.push( this.current ); this.restore( this.undoStack.pop() ); this.lastTyping = 0; } }
        redo() { if ( this.redoStack.length ) { this.undoStack.push( this.current ); this.restore( this.redoStack.pop() ); this.lastTyping = 0; } }
        get dirty() { return this.editable && this.current !== this.saved; }

        updateBar() {
            if ( !this.editable ) { return; }
            const btn = a => this.bar.querySelector( "[data-act=\"" + a + "\"]" );
            btn( "undo" ).disabled = !this.undoStack.length;
            btn( "redo" ).disabled = !this.redoStack.length;
            btn( "delete" ).disabled = !this.meta.id || !this.meta.mine;
            btn( "save" ).classList.toggle( "is-dirty", this.dirty );
            this.bar.querySelectorAll( "[data-act=\"view\"]" ).forEach( b => b.setAttribute( "aria-pressed", String( ( b.dataset.view === "names" ) === this.names ) ) );
            this.el.classList.toggle( "is-dirty", this.dirty );
            if ( this.dirty ) { this.setStatus( this.meta.mine ? "Unsaved changes" : "Unsaved changes. Save as to keep them." ); }
            else if ( this.statusText === "Unsaved changes" || /^Unsaved/.test( this.statusText || "" ) ) { this.setStatus( "" ); }
        }
        setStatus( text ) {
            this.statusText = text;
            if ( this.status ) { this.status.textContent = text || ""; this.status.classList.toggle( "is-dirty", /^Unsaved/.test( text || "" ) ); }
        }

        confirmDiscard( ok, cancel ) {
            if ( !this.dirty ) { ok(); return; }
            apex.message.confirm( "This template has unsaved changes. Discard them?", yes => { if ( yes ) { ok(); } else if ( cancel ) { cancel(); } } );
        }

        // ---------- blocks ----------
        block( id ) { return this.tpl.blocks.find( b => b.id === id ); }
        index( id ) { return this.tpl.blocks.findIndex( b => b.id === id ); }

        addBlock( type, at ) {
            const b = Object.assign( { id: uid(), type }, BLOCKS[ type ].make( this.data ) );
            if ( at === undefined || at === null ) {
                const i = this.selected ? this.index( this.selected ) : -1;
                at = i < 0 ? this.tpl.blocks.length : i + 1;
            }
            this.tpl.blocks.splice( at, 0, b );
            this.commit();
            this.selected = b.id;
            this.render();
            this.setTab( "block" );
            this.scrollToBlock( b.id );
            this.announce( BLOCKS[ type ].label + " added" );
        }

        moveBlock( id, at ) {
            const i = this.index( id );
            if ( i < 0 ) { return; }
            const [ b ] = this.tpl.blocks.splice( i, 1 );
            this.tpl.blocks.splice( at > i ? at - 1 : at, 0, b );
            this.commit();
            this.selected = id;
            this.render();
            this.renderProps();
        }

        blockAction( act, id ) {
            const i = this.index( id );
            if ( i < 0 ) { return; }
            if ( act === "up" && i > 0 ) { this.moveBlock( id, i - 1 ); this.focusBlock( id ); }
            else if ( act === "down" && i < this.tpl.blocks.length - 1 ) { this.moveBlock( id, i + 2 ); this.focusBlock( id ); }
            else if ( act === "copy" ) {
                const b = clone( this.tpl.blocks[ i ] );
                b.id = uid();
                this.tpl.blocks.splice( i + 1, 0, b );
                this.commit();
                this.selected = b.id;
                this.render();
                this.renderProps();
                this.announce( "Block duplicated" );
            } else if ( act === "delete" ) {
                const label = BLOCKS[ this.tpl.blocks[ i ].type ].label;
                this.tpl.blocks.splice( i, 1 );
                this.commit();
                const next = this.tpl.blocks[ Math.min( i, this.tpl.blocks.length - 1 ) ];
                this.selected = next ? next.id : null;
                this.render();
                this.renderProps();
                if ( next ) { this.focusBlock( next.id ); }
                this.announce( label + " deleted. Ctrl+Z brings it back." );
            }
        }

        select( id ) {
            if ( this.selected === id ) { return; }
            this.selected = id;
            this.pagesEl.querySelectorAll( ".tx-Rd-slot" ).forEach( s => s.classList.toggle( "is-selected", s.dataset.id === id ) );
            this.render();
            this.setTab( id ? "block" : "report" );
        }

        focusBlock( id ) {
            const s = this.pagesEl.querySelector( ".tx-Rd-slot[data-id=\"" + id + "\"] .tx-Rd-block" );
            if ( s ) { s.focus(); }
        }
        focusProps() { const f = this.props.querySelector( "input, textarea, select" ); if ( f ) { f.focus(); } }
        scrollToBlock( id ) {
            const s = this.pagesEl.querySelector( ".tx-Rd-slot[data-id=\"" + id + "\"]" );
            if ( s ) { s.scrollIntoView( { block: "nearest", behavior: "smooth" } ); }
        }
        announce( text ) {
            this.setStatus( text );
            clearTimeout( this.announceTimer );
            this.announceTimer = setTimeout( () => { if ( this.statusText === text ) { this.setStatus( "" ); this.updateBar(); } }, 2600 );
        }

        // ---------- drag and drop on the page ----------
        showDrop( y, target ) {
            const sheet = ( target.closest && target.closest( ".tx-Rd-sheet" ) ) || this.pagesEl.querySelector( ".tx-Rd-sheet" );
            if ( !sheet ) { return; }
            const slots = [ ...sheet.querySelectorAll( ".tx-Rd-slot" ) ], area = sheet.querySelector( ".tx-Rd-blocks" );
            let idx = slots.length, lineY;
            for ( let i = 0; i < slots.length; i++ ) {
                const r = slots[ i ].getBoundingClientRect();
                if ( y < r.top + r.height / 2 ) { idx = i; lineY = r.top - 2; break; }
            }
            const ar = area.getBoundingClientRect(), sr = this.stage.getBoundingClientRect();
            if ( lineY === undefined ) { lineY = slots.length ? slots[ slots.length - 1 ].getBoundingClientRect().bottom + 2 : ar.top + 8; }
            this.dropIndex = idx;
            Object.assign( this.dropEl.style, { top: ( lineY - sr.top + this.stage.scrollTop ) + "px", left: ( ar.left - sr.left + this.stage.scrollLeft ) + "px", width: ar.width + "px" } );
            this.dropEl.hidden = false;
        }
        hideDrop() { this.dropEl.hidden = true; this.dropIndex = null; }

        // ---------- rendering ----------
        fit() {
            const paper = this.pagesEl.querySelector( ".tx-Rd-paper" );
            if ( !paper ) { return; }
            const [ W ] = pageSize( this.tpl ), avail = this.stage.clientWidth - 40;
            this.pagesEl.style.zoom = Math.max( 0.3, Math.min( 1, avail / ( W * PX ) ) ).toFixed( 3 );
        }

        render() {
            const t = this.tpl, [ W, H ] = pageSize( t ), m = +t.page.margin || 40;
            if ( !this.editable && !this.templates.length ) {
                this.pagesEl.innerHTML = "<div class=\"tx-Rd-placeholder\"><span class=\"fa fa-file-text-o\" aria-hidden=\"true\"></span><p>No report templates yet. Templates designed in the Report Designer appear here.</p></div>";
                this.bar.querySelectorAll( "[data-act=\"download\"]" ).forEach( b => { b.disabled = true; } );
                return;
            }
            const model = build( t, this.data, { limit: 40, maxGroups: this.editable ? 2 : 3, name: this.meta.name, user: this.cfg.user } );
            const names = this.names, ed = this.editable;
            const vars = "--rd-w:" + W + "pt;--rd-h:" + H + "pt;--rd-m:" + m + "pt;--rd-fs:" + ( +t.style.fontSize || 10 ) + "pt;--rd-accent:#" + hex6( t.style.accent, "0F766E" )
                + ";--rd-tint:#" + mix( hex6( t.style.accent, "0F766E" ), "FFFFFF", 0.88 ) + ";--rd-text:#" + hex6( t.style.text );
            const h = t.header, ft = t.footer;
            let html = "";
            model.groups.forEach( ( g, gi ) => {
                const ctx = g.ctx, first = gi === 0;
                const head = h.left || h.right || h.logo
                    ? "<div class=\"tx-Rd-pHead" + ( h.line ? " has-line" : "" ) + "\"" + ( ed ? " title=\"Header: edit it in the Report tab\"" : "" ) + ">"
                      + ( h.logo ? "<img class=\"tx-Rd-logo\" alt=\"\" src=\"" + esc( h.logo ) + "\">" : "" )
                      + "<div class=\"tx-Rd-pHeadL\">" + resolveHtml( h.left, ctx, names, 1 ) + "</div><div class=\"tx-Rd-pHeadR\">" + resolveHtml( h.right, ctx, names, 1 ) + "</div></div>"
                    : ( ed && first ? "<div class=\"tx-Rd-pHead is-empty\" title=\"Add a header in the Report tab\">Header</div>" : "" );
                const foot = ft.left || ft.right
                    ? "<div class=\"tx-Rd-pFoot\"><div>" + resolveHtml( ft.left, ctx, names, 1 ) + "</div><div>" + resolveHtml( ft.right, ctx, names, 1 ) + "</div></div>"
                    : "";
                const blocks = g.els.map( el => {
                    const label = BLOCKS[ el.type ].label, sel = el.id === this.selected;
                    const inner = blockHtml( el, ctx, names );
                    if ( !ed ) { return "<div class=\"tx-Rd-slot is-static\">" + inner + "</div>"; }
                    return "<div class=\"tx-Rd-slot tx-Rd-slot--" + el.type + ( sel ? " is-selected" : "" ) + "\" data-id=\"" + esc( el.id ) + "\">"
                        + "<div class=\"tx-Rd-block\"" + ( first ? " tabindex=\"0\"" : "" ) + " aria-label=\"" + esc( label + " block" + ( sel ? ", selected" : "" ) ) + "\">" + inner + "</div>"
                        + ( sel && first ? "<div class=\"tx-Rd-blockBar\" role=\"toolbar\" aria-label=\"" + esc( label ) + " block\">"
                            + "<span class=\"tx-Rd-blockName\">" + esc( label ) + "</span>"
                            + "<span class=\"tx-Rd-handle\" draggable=\"true\" data-block-act=\"move\" title=\"Drag to move\" aria-hidden=\"true\">" + ICON( "fa-arrows" ) + "</span>"
                            + "<button type=\"button\" data-block-act=\"up\" title=\"Move up (Alt+Up)\" aria-label=\"Move up\">" + ICON( "fa-arrow-up" ) + "</button>"
                            + "<button type=\"button\" data-block-act=\"down\" title=\"Move down (Alt+Down)\" aria-label=\"Move down\">" + ICON( "fa-arrow-down" ) + "</button>"
                            + "<button type=\"button\" data-block-act=\"copy\" title=\"Duplicate\" aria-label=\"Duplicate\">" + ICON( "fa-clone" ) + "</button>"
                            + "<button type=\"button\" data-block-act=\"delete\" title=\"Delete (Del)\" aria-label=\"Delete\">" + ICON( "fa-trash-o" ) + "</button>"
                            + "</div>" : "" )
                        + "</div>";
                } ).join( "" );
                const empty = !t.blocks.length && ed
                    ? "<div class=\"tx-Rd-dropHint\"><span class=\"fa fa-hand-pointer-o\" aria-hidden=\"true\"></span><strong>Start your report</strong><span>Drag a block from the left onto the page, or click one to add it.</span></div>" : "";
                const label = model.groupBy ? "<div class=\"tx-Rd-sheetLabel\">" + esc( labelOf( model.groupBy ) ) + ": " + esc( model.f.value( g.key, this.data.types[ model.groupBy ], "" ) )
                    + ( first && t.settings.groupPageBreak ? "<span> (each on a new page)</span>" : "" ) + "</div>" : "";
                html += "<section class=\"tx-Rd-sheet\">" + label + "<div class=\"tx-Rd-paper\" style=\"" + vars + "\">" + head + "<div class=\"tx-Rd-blocks\">" + blocks + empty + "</div>" + foot + "</div></section>";
            } );
            if ( model.moreGroups > 0 ) {
                html += "<p class=\"tx-Rd-moreGroups\">" + esc( model.moreGroups ) + " more " + ( model.moreGroups === 1 ? "section" : "sections" ) + " in the downloaded report</p>";
            }
            if ( this.data.truncated ) {
                html += "<p class=\"tx-Rd-moreGroups\">The data has more rows than the Maximum Rows setting; only the first rows are used.</p>";
            }
            if ( !this.data.rows.length && this.data.columns.length ) {
                html = "<p class=\"tx-Rd-notice\">" + esc( this.cfg.noDataFound || "No data found. The report shows its layout without rows." ) + "</p>" + html;
            }
            const keep = document.activeElement && this.pagesEl.contains( document.activeElement ) && document.activeElement.closest( ".tx-Rd-slot" );
            const keepId = keep && keep.dataset.id;
            this.pagesEl.innerHTML = html;
            this.fit();
            if ( keepId ) { this.focusBlock( keepId ); }
            this.bar.querySelectorAll( "[data-act=\"download\"]" ).forEach( b => { b.disabled = !this.data.columns.length; } );
        }

        renderFields( filter ) {
            if ( !this.fieldList ) { return; }
            const q = String( filter || "" ).trim().toUpperCase();
            const icon = { number: "fa-hashtag", date: "fa-calendar", text: "fa-font" };
            const cols = this.data.columns.filter( c => !q || c.name.indexOf( q ) > -1 || labelOf( c.name ).toUpperCase().indexOf( q ) > -1 );
            this.fieldList.innerHTML = cols.length
                ? cols.map( c => "<li><div role=\"button\" tabindex=\"0\" class=\"tx-Rd-field\" draggable=\"true\" data-field=\"" + esc( c.name ) + "\" title=\"Insert {" + esc( c.name ) + "}\">"
                    + "<span class=\"fa " + ( icon[ c.type ] || "fa-font" ) + "\" aria-hidden=\"true\"></span><span class=\"tx-Rd-fieldName\">" + esc( c.name ) + "</span>"
                    + "<span class=\"tx-Rd-fieldType\">" + esc( c.type ) + "</span></div></li>" ).join( "" )
                : "<li class=\"tx-Rd-sideHint\">" + ( this.data.columns.length ? "No field matches." : "The region source has no columns yet." ) + "</li>";
        }

        // ---------- properties panel ----------
        setTab( tab, focusSection ) {
            if ( !this.editable ) { return; }
            this.tab = tab;
            this.el.querySelectorAll( ".tx-Rd-tabs [data-tab]" ).forEach( b => {
                const on = b.dataset.tab === tab;
                b.setAttribute( "aria-selected", String( on ) );
                b.tabIndex = on ? 0 : -1;
            } );
            this.props.setAttribute( "aria-labelledby", this.id + ( tab === "block" ? "_tabBlock" : "_tabReport" ) );
            this.renderProps();
            if ( focusSection ) {
                const d = this.props.querySelector( "details[data-section=\"" + focusSection + "\"]" );
                if ( d ) { d.open = true; d.scrollIntoView( { block: "nearest" } ); }
            }
        }

        renderProps() {
            if ( !this.editable ) { return; }
            const scroll = this.props.scrollTop;
            this.props.innerHTML = this.tab === "block" ? this.blockProps() : this.reportProps();
            this.props.scrollTop = scroll;
        }

        // control helpers; data-p: path in the block, data-t: path in the template, data-m: path in the meta data
        ctl( kind, label, path, value, extra ) {
            const id = this.id + "_p" + ( ++this.ctlSeq || ( this.ctlSeq = 1 ) ), x = extra || {};
            const attr = path.indexOf( "t:" ) === 0 ? "data-t=\"" + esc( path.slice( 2 ) ) + "\"" : path.indexOf( "m:" ) === 0 ? "data-m=\"" + esc( path.slice( 2 ) ) + "\"" : "data-p=\"" + esc( path ) + "\"";
            const ins = x.fields ? "<button type=\"button\" class=\"tx-Rd-insert\" data-insert=\"" + id + "\" title=\"Insert a field\" aria-label=\"Insert a field into " + esc( label ) + "\">{ }</button>" : "";
            const fieldsOk = x.fields ? " data-fields-ok" : "";
            let control;
            switch ( kind ) {
                case "textarea":
                    control = "<textarea id=\"" + id + "\" class=\"tx-Rd-input\" rows=\"" + ( x.rows || 4 ) + "\" " + attr + fieldsOk + ">" + esc( value ) + "</textarea>"; break;
                case "select":
                    control = "<select id=\"" + id + "\" class=\"tx-Rd-input\" " + attr + ( x.num ? " data-num" : "" ) + ">" + options( x.list, value ) + "</select>"; break;
                case "number":
                    control = "<input id=\"" + id + "\" type=\"number\" class=\"tx-Rd-input tx-Rd-input--num\" " + attr + " data-num value=\"" + esc( value ) + "\""
                        + ( x.min !== undefined ? " min=\"" + x.min + "\"" : "" ) + ( x.max !== undefined ? " max=\"" + x.max + "\"" : "" ) + ( x.step ? " step=\"" + x.step + "\"" : "" )
                        + ( x.placeholder ? " placeholder=\"" + esc( x.placeholder ) + "\"" : "" ) + ">" + ( x.unit ? "<span class=\"tx-Rd-unit\">" + esc( x.unit ) + "</span>" : "" ); break;
                case "check":
                    return "<label class=\"tx-Rd-check\"><input type=\"checkbox\" " + attr + ( value ? " checked" : "" ) + "><span>" + esc( label ) + "</span></label>";
                case "seg":
                    return "<div class=\"tx-Rd-field\"><span class=\"tx-Rd-label\" id=\"" + id + "\">" + esc( label ) + "</span><div class=\"tx-Rd-seg tx-Rd-seg--full\" role=\"group\" aria-labelledby=\"" + id + "\">"
                        + x.list.map( o => "<button type=\"button\" " + attr + " data-value=\"" + esc( o[ 0 ] ) + "\"" + ( x.num ? " data-num" : "" ) + " aria-pressed=\"" + ( String( o[ 0 ] ) === String( value ) ) + "\""
                            + ( o[ 2 ] ? " title=\"" + esc( o[ 1 ] ) + "\" aria-label=\"" + esc( o[ 1 ] ) + "\">" + ICON( o[ 2 ] ) : ">" + esc( o[ 1 ] ) ) + "</button>" ).join( "" ) + "</div></div>";
                case "color":
                    control = "<span class=\"tx-Rd-colorRow\"><input id=\"" + id + "\" type=\"color\" class=\"tx-Rd-color\" " + attr + " value=\"#" + hex6( value || x.auto || "#1f2937" ).toLowerCase() + "\">"
                        + ( x.auto !== undefined ? "<button type=\"button\" class=\"tx-Rd-link\" data-reset=\"" + esc( path ) + "\"" + ( value ? "" : " disabled" ) + ">" + ( value ? "Use automatic color" : "Automatic" ) + "</button>" : "" ) + "</span>"; break;
                default:
                    control = "<input id=\"" + id + "\" type=\"text\" class=\"tx-Rd-input\" " + attr + fieldsOk + " value=\"" + esc( value ) + "\"" + ( x.placeholder ? " placeholder=\"" + esc( x.placeholder ) + "\"" : "" ) + ">";
            }
            return "<div class=\"tx-Rd-field\"><div class=\"tx-Rd-labelRow\"><label class=\"tx-Rd-label\" for=\"" + id + "\">" + esc( label ) + "</label>" + ins + "</div>"
                + "<div class=\"tx-Rd-control\">" + control + "</div>" + ( x.help ? "<p class=\"tx-Rd-help\">" + x.help + "</p>" : "" ) + "</div>";
        }

        colOptions( blank ) {
            return ( blank ? [ [ "", blank ] ] : [] ).concat( this.data.columns.map( c => [ c.name, labelOf( c.name ) + " (" + c.name + ")" ] ) );
        }

        blockProps() {
            const b = this.block( this.selected );
            if ( !b ) {
                return "<div class=\"tx-Rd-propsEmpty\"><span class=\"fa fa-mouse-pointer\" aria-hidden=\"true\"></span><p>Click a block on the page to change it, or add one from the left.</p></div>";
            }
            const B = BLOCKS[ b.type ], align = [ [ "left", "Left", "fa-align-left" ], [ "center", "Center", "fa-align-center" ], [ "right", "Right", "fa-align-right" ] ];
            let html = "<div class=\"tx-Rd-propsHead\"><span class=\"fa " + B.icon + "\" aria-hidden=\"true\"></span><strong>" + esc( B.label ) + "</strong>"
                + "<span class=\"tx-Rd-propsActions\"><button type=\"button\" class=\"tx-Rd-btn tx-Rd-btn--icon\" data-pact=\"copy\" title=\"Duplicate\" aria-label=\"Duplicate block\">" + ICON( "fa-clone" ) + "</button>"
                + "<button type=\"button\" class=\"tx-Rd-btn tx-Rd-btn--icon\" data-pact=\"delete\" title=\"Delete\" aria-label=\"Delete block\">" + ICON( "fa-trash-o" ) + "</button></span></div>";
            switch ( b.type ) {
                case "heading":
                    html += this.ctl( "text", "Text", "text", b.text, { fields: true } )
                        + this.ctl( "seg", "Style", "level", b.level, { list: [ [ 1, "Title" ], [ 2, "Heading" ], [ 3, "Subheading" ] ], num: true } )
                        + this.ctl( "seg", "Alignment", "align", b.align || "left", { list: align } )
                        + this.ctl( "color", "Color", "color", b.color, { auto: +b.level === 1 ? this.tpl.style.text : this.tpl.style.accent } );
                    break;
                case "text":
                    html += this.ctl( "textarea", "Text", "text", b.text, { fields: true, rows: 6, help: "Fields such as {CUSTOMER} show the value of the first row; totals such as {SUM:AMOUNT|number} use all rows." } )
                        + "<div class=\"tx-Rd-row\">" + this.ctl( "number", "Font size", "size", +b.size || "", { min: 6, max: 40, placeholder: String( this.tpl.style.fontSize ), unit: "pt" } )
                        + "<div class=\"tx-Rd-field\"><span class=\"tx-Rd-label\">Style</span><div class=\"tx-Rd-checks\">" + this.ctl( "check", "Bold", "bold", b.bold ) + this.ctl( "check", "Italic", "italic", b.italic ) + "</div></div></div>"
                        + this.ctl( "seg", "Alignment", "align", b.align || "left", { list: align.concat( [] ) } )
                        + this.ctl( "color", "Color", "color", b.color, { auto: this.tpl.style.text } );
                    break;
                case "fields":
                    html += "<div class=\"tx-Rd-field\"><span class=\"tx-Rd-label\">Fields</span><ol class=\"tx-Rd-items\">"
                        + ( b.items || [] ).map( ( it, i ) => "<li class=\"tx-Rd-item\">"
                            + "<div class=\"tx-Rd-itemRow\"><input type=\"text\" class=\"tx-Rd-input\" data-p=\"items." + i + ".label\" value=\"" + esc( it.label ) + "\" aria-label=\"Label " + ( i + 1 ) + "\" placeholder=\"Label\">"
                            + "<button type=\"button\" class=\"tx-Rd-btn tx-Rd-btn--icon tx-Rd-btn--sm\" data-pact=\"itemUp\" data-i=\"" + i + "\" title=\"Move up\" aria-label=\"Move up\"" + ( i ? "" : " disabled" ) + ">" + ICON( "fa-arrow-up" ) + "</button>"
                            + "<button type=\"button\" class=\"tx-Rd-btn tx-Rd-btn--icon tx-Rd-btn--sm\" data-pact=\"itemDel\" data-i=\"" + i + "\" title=\"Remove\" aria-label=\"Remove field " + ( i + 1 ) + "\">" + ICON( "fa-times" ) + "</button></div>"
                            + "<div class=\"tx-Rd-itemRow\"><input type=\"text\" id=\"" + this.id + "_iv" + i + "\" class=\"tx-Rd-input\" data-p=\"items." + i + ".value\" data-fields-ok value=\"" + esc( it.value ) + "\" aria-label=\"Value " + ( i + 1 ) + "\" placeholder=\"Value, for example {CUSTOMER}\">"
                            + "<button type=\"button\" class=\"tx-Rd-insert\" data-insert=\"" + this.id + "_iv" + i + "\" title=\"Insert a field\" aria-label=\"Insert a field into value " + ( i + 1 ) + "\">{ }</button></div></li>" ).join( "" )
                        + "</ol><div class=\"tx-Rd-addRow\"><select class=\"tx-Rd-input\" data-pact=\"itemAdd\" aria-label=\"Add a field\">" + options( this.colOptions( "Add a field..." ), "" ) + "<option value=\"*\">Empty row</option></select></div></div>"
                        + this.ctl( "seg", "Columns", "columns", b.columns || 2, { list: [ [ 1, "1" ], [ 2, "2" ], [ 3, "3" ] ], num: true } )
                        + this.ctl( "seg", "Position", "align", b.align === "right" ? "right" : "left", { list: [ [ "left", "Full width" ], [ "right", "Right (totals)" ] ] } );
                    break;
                case "table":
                    html += this.tableProps( b );
                    break;
                case "image":
                    html += "<div class=\"tx-Rd-field\"><span class=\"tx-Rd-label\">Image</span>"
                        + ( b.src ? "<img class=\"tx-Rd-thumb\" alt=\"\" src=\"" + esc( b.src ) + "\">" : "" )
                        + "<div class=\"tx-Rd-btnRow\"><label class=\"tx-Rd-btn tx-Rd-file\">" + ICON( "fa-upload", b.src ? "Replace image" : "Choose image" )
                        + "<input type=\"file\" accept=\"image/*\" data-image=\"p:src\" class=\"tx-Rd-srOnly\"></label>"
                        + ( b.src ? "<button type=\"button\" class=\"tx-Rd-btn\" data-pact=\"imgDel\">" + ICON( "fa-times", "Remove" ) + "</button>" : "" ) + "</div>"
                        + "<p class=\"tx-Rd-help\">PNG, JPEG, SVG or GIF; saved inside the template.</p></div>"
                        + this.ctl( "number", "Width", "width", b.width || 140, { min: 16, max: 800, unit: "pt" } )
                        + this.ctl( "seg", "Alignment", "align", b.align || "left", { list: align } );
                    break;
                case "divider":
                    html += this.ctl( "seg", "Thickness", "thickness", b.thickness || 1, { list: [ [ 0.5, "Hairline" ], [ 1, "Thin" ], [ 2, "Medium" ], [ 3, "Thick" ] ], num: true } )
                        + this.ctl( "color", "Color", "color", b.color, { auto: LINE } );
                    break;
                case "spacer":
                    html += this.ctl( "number", "Height", "height", b.height || 16, { min: 2, max: 400, unit: "pt" } );
                    break;
                case "pagebreak":
                    html += "<p class=\"tx-Rd-help\">The block after this one starts on a new page in PDF and Word, and after a page break in Excel.</p>";
                    break;
            }
            return html;
        }

        tableProps( b ) {
            const align = [ [ "left", "Left", "fa-align-left" ], [ "center", "Center", "fa-align-center" ], [ "right", "Right", "fa-align-right" ] ];
            const cols = ( b.columns || [] ).map( ( c, i ) => {
                const unknown = c.field && !this.data.has( c.field );
                return "<li><details class=\"tx-Rd-col\"" + ( this.openCol === i ? " open" : "" ) + " data-col=\"" + i + "\"><summary>"
                    + "<span class=\"tx-Rd-colName\">" + esc( c.label || c.field || "Column" ) + "<small" + ( unknown ? " class=\"is-bad\"" : "" ) + ">" + esc( c.field || "no field" ) + ( unknown ? ", not in the data" : "" )
                    + ( c.total ? ", " + esc( TOTAL_NAMES[ c.total ] ) : "" ) + "</small></span>"
                    + "<span class=\"tx-Rd-colBtns\"><button type=\"button\" class=\"tx-Rd-btn tx-Rd-btn--icon tx-Rd-btn--sm\" data-pact=\"colUp\" data-i=\"" + i + "\" title=\"Move up\" aria-label=\"Move column up\"" + ( i ? "" : " disabled" ) + ">" + ICON( "fa-arrow-up" ) + "</button>"
                    + "<button type=\"button\" class=\"tx-Rd-btn tx-Rd-btn--icon tx-Rd-btn--sm\" data-pact=\"colDel\" data-i=\"" + i + "\" title=\"Remove column\" aria-label=\"Remove column " + esc( c.label || c.field ) + "\">" + ICON( "fa-times" ) + "</button></span>"
                    + "</summary><div class=\"tx-Rd-colBody\">"
                    + this.ctl( "select", "Field", "columns." + i + ".field", c.field, { list: this.colOptions( "Choose a field" ) } )
                    + this.ctl( "text", "Heading", "columns." + i + ".label", c.label )
                    + "<div class=\"tx-Rd-row\">" + this.ctl( "select", "Format", "columns." + i + ".format", c.format, { list: FORMATS } )
                    + this.ctl( "select", "Total", "columns." + i + ".total", c.total, { list: TOTALS } ) + "</div>"
                    + "<div class=\"tx-Rd-row\">" + this.ctl( "select", "Width", "columns." + i + ".width", c.width || 1, { list: WIDTHS, num: true } ) + "</div>"
                    + this.ctl( "seg", "Alignment", "columns." + i + ".align", c.align || "left", { list: align } )
                    + "</div></details></li>";
            } ).join( "" );
            return "<div class=\"tx-Rd-field\"><span class=\"tx-Rd-label\">Columns</span><ol class=\"tx-Rd-cols\">" + cols + "</ol>"
                + "<div class=\"tx-Rd-addRow\"><select class=\"tx-Rd-input\" data-pact=\"colAdd\" aria-label=\"Add a column\">" + options( this.colOptions( "Add a column..." ), "" ) + "</select>"
                + "<button type=\"button\" class=\"tx-Rd-btn\" data-pact=\"colAll\" title=\"Add every field that is not in the table yet\">Add all</button></div></div>"
                + this.ctl( "select", "Group rows by", "groupBy", b.groupBy || "", { list: this.colOptions( "No groups" ), help: "A heading row for each value, with subtotals." } )
                + ( b.groupBy ? this.ctl( "check", "Subtotal for each group", "subtotals", b.subtotals !== false ) : "" )
                + this.ctl( "check", "Totals row", "totals", b.totals !== false )
                + this.ctl( "check", "Striped rows", "striped", b.striped !== false )
                + this.ctl( "seg", "Lines", "border", b.border || "lines", { list: [ [ "lines", "Rows" ], [ "grid", "Grid" ], [ "none", "None" ] ] } )
                + this.ctl( "number", "Font size", "size", +b.size || "", { min: 6, max: 20, step: 0.5, placeholder: String( ( +this.tpl.style.fontSize || 10 ) - 0.5 ), unit: "pt" } );
        }

        reportProps() {
            const t = this.tpl, m = this.meta, today = new Date(), f = new Formatter( t );
            const section = ( key, title, body, open ) => "<details class=\"tx-Rd-section\" data-section=\"" + key + "\"" + ( open ? " open" : "" ) + "><summary>" + esc( title ) + "</summary><div class=\"tx-Rd-sectionBody\">" + body + "</div></details>";
            const dateList = DATE_FORMATS.map( d => [ d, f.date.call( { s: { dateFormat: d }, locale: f.locale }, today ) ] );
            const sizeList = Object.keys( PAGE_SIZES ).map( k => [ k, k ] );
            return section( "template", "Template", this.ctl( "text", "Name", "m:name", m.name, { placeholder: "Report name" } )
                    + this.ctl( "textarea", "Description", "m:description", m.description, { rows: 2 } )
                    + ( this.cfg.sharing ? this.ctl( "check", "Share with everyone (read only for others)", "m:shared", m.shared ) : "" ), true )
                + section( "data", "Sections", this.ctl( "select", "One section for each", "t:settings.groupBy", t.settings.groupBy, { list: this.colOptions( "No sections (one report)" ),
                        help: "Repeats every block for each value, for example one invoice per order. Fields and totals then use that section's rows." } )
                    + ( t.settings.groupBy ? this.ctl( "check", "Start each section on a new page", "t:settings.groupPageBreak", t.settings.groupPageBreak !== false )
                        + this.ctl( "check", "Excel: one sheet per section", "t:settings.sheetPerGroup", t.settings.sheetPerGroup !== false ) : "" ), !!t.settings.groupBy )
                + section( "page", "Page", "<div class=\"tx-Rd-row\">" + this.ctl( "select", "Size", "t:page.size", t.page.size, { list: sizeList } )
                    + this.ctl( "select", "Margins", "t:page.margin", t.page.margin, { list: [ [ 24, "Narrow" ], [ 40, "Normal" ], [ 60, "Wide" ] ], num: true } ) + "</div>"
                    + this.ctl( "seg", "Orientation", "t:page.orientation", t.page.orientation, { list: [ [ "portrait", "Portrait" ], [ "landscape", "Landscape" ] ] } ) )
                + section( "style", "Style", "<div class=\"tx-Rd-row\">" + this.ctl( "number", "Font size", "t:style.fontSize", t.style.fontSize, { min: 7, max: 16, step: 0.5, unit: "pt" } ) + "</div>"
                    + "<div class=\"tx-Rd-row\">" + this.ctl( "color", "Accent color", "t:style.accent", t.style.accent ) + this.ctl( "color", "Text color", "t:style.text", t.style.text ) + "</div>" )
                + section( "header", "Header and footer",
                    "<div class=\"tx-Rd-field\"><span class=\"tx-Rd-label\">Logo</span>" + ( t.header.logo ? "<img class=\"tx-Rd-thumb tx-Rd-thumb--logo\" alt=\"\" src=\"" + esc( t.header.logo ) + "\">" : "" )
                    + "<div class=\"tx-Rd-btnRow\"><label class=\"tx-Rd-btn tx-Rd-file\">" + ICON( "fa-upload", t.header.logo ? "Replace logo" : "Choose logo" )
                    + "<input type=\"file\" accept=\"image/*\" data-image=\"t:header.logo\" class=\"tx-Rd-srOnly\"></label>"
                    + ( t.header.logo ? "<button type=\"button\" class=\"tx-Rd-btn\" data-pact=\"logoDel\">" + ICON( "fa-times", "Remove" ) + "</button>" : "" ) + "</div></div>"
                    + this.ctl( "text", "Header left", "t:header.left", t.header.left, { fields: true, placeholder: "For example your company name" } )
                    + this.ctl( "text", "Header right", "t:header.right", t.header.right, { fields: true, placeholder: "For example {TODAY}" } )
                    + this.ctl( "check", "Line below the header", "t:header.line", t.header.line )
                    + this.ctl( "text", "Footer left", "t:footer.left", t.footer.left, { fields: true } )
                    + this.ctl( "text", "Footer right", "t:footer.right", t.footer.right, { fields: true, help: "{PAGE} and {PAGES} give the page number and the page count." } ), !!( t.header.left || t.header.right || t.header.logo ) )
                + section( "formats", "Formats and file name", this.ctl( "select", "Date format", "t:settings.dateFormat", t.settings.dateFormat, { list: dateList } )
                    + this.ctl( "text", "Currency symbol", "t:settings.currency", t.settings.currency, { placeholder: "$" } )
                    + this.ctl( "text", "File name", "t:settings.fileName", t.settings.fileName, { fields: true, help: "Without extension. {REPORT} is the template name." } ) );
        }

        propInput( e, final ) {
            const el = e.target;
            if ( el.matches( "[data-image]" ) ) { if ( final ) { this.pickImage( el ); } return; }
            if ( el.matches( "[data-pact]" ) ) { if ( final ) { this.propSelectAction( el ); } return; }
            const isText = el.matches( "input[type=text], input[type=search], textarea" );
            if ( !final && !isText && el.type !== "number" && el.type !== "color" ) { return; }
            if ( final && isText ) { return; }   // typed text was applied on input
            let v = el.type === "checkbox" ? el.checked : el.value;
            if ( el.hasAttribute( "data-num" ) ) { v = v === "" ? 0 : Number( v ); }
            this.apply( el.dataset, v, isText || el.type === "number" || el.type === "color" );
            // a different field or structure: redraw the panel
            if ( final && ( el.tagName === "SELECT" || el.type === "checkbox" ) ) {
                if ( el.dataset.p && /^columns\.\d+\.field$/.test( el.dataset.p ) ) {
                    const i = +el.dataset.p.split( "." )[ 1 ], c = this.block( this.selected ).columns[ i ];
                    const col = this.data.columns.find( x => x.name === v );
                    if ( col ) { Object.assign( c, colDef( col ), { width: c.width } ); this.commit(); this.render(); }
                    this.openCol = i;
                }
                this.renderProps();
            }
        }

        apply( ds, v, typing ) {
            if ( ds.p !== undefined ) {
                const b = this.block( this.selected );
                if ( !b ) { return; }
                setPath( b, ds.p, v );
            } else if ( ds.t !== undefined ) {
                setPath( this.tpl, ds.t, v );
            } else if ( ds.m !== undefined ) {
                this.meta[ ds.m ] = v;
                if ( ds.m === "name" ) { this.fillTemplateSelect(); }
            }
            this.commit( typing );
            clearTimeout( this.renderTimer );
            this.renderTimer = setTimeout( () => this.render(), typing ? 140 : 0 );
        }

        propClick( e ) {
            const seg = e.target.closest( ".tx-Rd-seg button[data-value]" );
            if ( seg ) {
                let v = seg.dataset.value;
                if ( seg.hasAttribute( "data-num" ) ) { v = Number( v ); }
                this.apply( seg.dataset, v, false );
                seg.parentElement.querySelectorAll( "button" ).forEach( b => b.setAttribute( "aria-pressed", String( b === seg ) ) );
                if ( seg.dataset.p === "level" || seg.dataset.t === "page.orientation" ) { this.renderProps(); }
                return;
            }
            const reset = e.target.closest( "[data-reset]" );
            if ( reset ) { const p = reset.dataset.reset; this.apply( p.indexOf( "t:" ) === 0 ? { t: p.slice( 2 ) } : { p }, "", false ); this.renderProps(); return; }
            const ins = e.target.closest( "[data-insert]" );
            if ( ins ) { e.preventDefault(); this.openPop( ins, document.getElementById( ins.dataset.insert ) ); return; }
            const a = e.target.closest( "button[data-pact]" );
            if ( !a ) { return; }
            e.preventDefault();
            e.stopPropagation();
            const b = this.block( this.selected ), i = +a.dataset.i;
            switch ( a.dataset.pact ) {
                case "copy": case "delete": this.blockAction( a.dataset.pact, this.selected ); return;
                case "itemUp": b.items.splice( i - 1, 0, b.items.splice( i, 1 )[ 0 ] ); break;
                case "itemDel": b.items.splice( i, 1 ); break;
                case "colUp": b.columns.splice( i - 1, 0, b.columns.splice( i, 1 )[ 0 ] ); this.openCol = i - 1; break;
                case "colDel": b.columns.splice( i, 1 ); this.openCol = null; break;
                case "colAll": this.data.columns.filter( c => !b.columns.some( x => x.field === c.name ) ).forEach( c => b.columns.push( colDef( c ) ) ); break;
                case "imgDel": b.src = ""; break;
                case "logoDel": this.tpl.header.logo = null; break;
                default: return;
            }
            this.commit();
            this.render();
            this.renderProps();
        }

        propSelectAction( sel ) {
            const b = this.block( this.selected ), v = sel.value;
            if ( !b || !v ) { return; }
            if ( sel.dataset.pact === "colAdd" ) {
                const col = this.data.columns.find( c => c.name === v );
                if ( col ) { b.columns.push( colDef( col ) ); this.openCol = b.columns.length - 1; }
            } else if ( sel.dataset.pact === "itemAdd" ) {
                b.items = b.items || [];
                b.items.push( v === "*" ? { label: "", value: "" } : { label: labelOf( v ), value: "{" + v + "}" } );
            }
            this.commit();
            this.render();
            this.renderProps();
        }

        pickImage( input ) {
            const file = input.files && input.files[ 0 ];
            if ( !file ) { return; }
            readImage( file ).then( img => {
                const path = input.dataset.image;
                if ( path === "t:header.logo" ) { this.tpl.header.logo = img.src; this.tpl.header.logoRatio = img.ratio; }
                else { const b = this.block( this.selected ); b.src = img.src; b.ratio = img.ratio; }
                this.commit();
                this.render();
                this.renderProps();
            } ).catch( e => this.showError( e ) );
        }

        // ---------- insert a field ----------
        insertField( token, target ) {
            const el = target || ( this.lastInput && this.lastInput.isConnected ? this.lastInput : null );
            if ( !el ) {
                if ( navigator.clipboard ) { navigator.clipboard.writeText( token ).catch( () => {} ); }
                this.announce( "Copied " + token + ". Click into a text box to insert it there." );
                return;
            }
            // at the cursor of a text box the user was typing in; otherwise at the end
            const typed = el === this.lastInput, len = el.value.length;
            const s = typed ? el.selectionStart ?? len : len, e = typed ? el.selectionEnd ?? len : len;
            if ( s === len && len && !/\s$/.test( el.value ) ) { token = " " + token; }
            el.value = el.value.slice( 0, s ) + token + el.value.slice( e );
            el.focus();
            el.setSelectionRange( s + token.length, s + token.length );
            this.lastInput = el;
            el.dispatchEvent( new Event( "input", { bubbles: true } ) );
        }

        openPop( btn, input ) {
            if ( !input ) { return; }
            this.popTarget = input;
            const cols = this.data.columns, nums = cols.filter( c => c.type === "number" && !isKey( c.name ) );
            const item = ( token, label, hint ) => "<li role=\"option\" tabindex=\"-1\" data-token=\"" + esc( token ) + "\"><span>" + esc( label ) + "</span><code>" + esc( hint || token ) + "</code></li>";
            this.pop.innerHTML = "<input type=\"search\" class=\"tx-Rd-input\" placeholder=\"Find a field\" aria-label=\"Find a field\">"
                + "<ul role=\"listbox\" aria-label=\"Fields\">"
                + "<li class=\"tx-Rd-popGroup\" role=\"presentation\">Fields</li>" + cols.map( c => item( "{" + c.name + "}", labelOf( c.name ) ) ).join( "" )
                + ( nums.length ? "<li class=\"tx-Rd-popGroup\" role=\"presentation\">Totals</li>" + nums.map( c => item( "{SUM:" + c.name + "|number}", "Sum of " + labelOf( c.name ).toLowerCase() ) ).join( "" )
                    + item( "{COUNT}", "Number of rows" ) : "" )
                + "<li class=\"tx-Rd-popGroup\" role=\"presentation\">Other</li>"
                + item( "{TODAY}", "Today's date" ) + item( "{REPORT}", "Template name" ) + item( "{APP_USER}", "User name" )
                + ( this.tpl.settings.groupBy ? item( "{GROUP}", "Section value" ) : "" )
                + item( "{PAGE}", "Page number" ) + item( "{PAGES}", "Page count" )
                + "</ul>";
            const r = btn.getBoundingClientRect(), host = this.el.getBoundingClientRect();
            this.pop.hidden = false;
            const left = Math.min( r.right - host.left - this.pop.offsetWidth, host.width - this.pop.offsetWidth - 8 );
            Object.assign( this.pop.style, { top: ( r.bottom - host.top + 4 ) + "px", left: Math.max( 8, left ) + "px" } );
            const search = this.pop.querySelector( "input" ), list = this.pop.querySelector( "ul" );
            const visible = () => [ ...list.querySelectorAll( "[data-token]" ) ].filter( li => !li.hidden );
            search.focus();
            search.oninput = () => {
                const q = search.value.trim().toLowerCase();
                list.querySelectorAll( "[data-token]" ).forEach( li => { li.hidden = q && li.textContent.toLowerCase().indexOf( q ) < 0; } );
            };
            const choose = li => { this.closePop(); this.insertField( li.dataset.token, input ); };
            list.onclick = e => { const li = e.target.closest( "[data-token]" ); if ( li ) { choose( li ); } };
            this.pop.onkeydown = e => {
                const items = visible(), i = items.indexOf( document.activeElement );
                if ( e.key === "Escape" ) { e.preventDefault(); this.closePop(); btn.focus(); }
                else if ( e.key === "ArrowDown" ) { e.preventDefault(); ( items[ i + 1 ] || items[ 0 ] ).focus(); }
                else if ( e.key === "ArrowUp" ) { e.preventDefault(); ( i > 0 ? items[ i - 1 ] : search ).focus(); }
                else if ( e.key === "Enter" ) { e.preventDefault(); const li = i > -1 ? items[ i ] : items[ 0 ]; if ( li ) { choose( li ); } }
            };
        }
        closePop() { this.pop.hidden = true; this.pop.innerHTML = ""; }

        // ---------- dialogs ----------
        openDialog( html, onSubmit ) {
            const d = this.dialog;
            if ( d.open ) { return; }
            d.innerHTML = "<form method=\"dialog\" class=\"tx-Rd-dlg\">" + html + "</form>";
            const form = d.querySelector( "form" );
            form.addEventListener( "submit", e => {
                if ( e.submitter && e.submitter.value === "cancel" ) { return; }
                e.preventDefault();
                if ( onSubmit( form ) !== false ) { d.close(); }
            } );
            d.showModal();
            const f = d.querySelector( "[autofocus]" ) || d.querySelector( "input, select" );
            if ( f ) { f.focus(); if ( f.select ) { f.select(); } }
        }

        newDialog() {
            const keyCols = this.data.columns.filter( c => c.type !== "number" || isKey( c.name ) );
            const guess = ( () => {
                let best = null;
                this.data.columns.forEach( c => {
                    const n = new Set( this.data.rows.map( r => keyOf( r[ c.name ] ) ) ).size;
                    if ( n > 1 && n < this.data.rows.length && ( !best || n > best.n ) && ( c.type !== "number" || isKey( c.name ) ) ) { best = { c, n }; }
                } );
                return best ? best.c.name : ( keyCols[ 0 ] || this.data.columns[ 0 ] || {} ).name;
            } )();
            const card = ( k, icon, title, text, checked ) => "<label class=\"tx-Rd-card\"><input type=\"radio\" name=\"kind\" value=\"" + k + "\"" + ( checked ? " checked" : "" ) + ">"
                + "<span class=\"fa " + icon + "\" aria-hidden=\"true\"></span><strong>" + title + "</strong><span>" + text + "</span></label>";
            this.confirmDiscard( () => { this.openDialog(
                "<h2 id=\"" + this.id + "_dlgTitle\">New template</h2>"
                + "<label class=\"tx-Rd-label\" for=\"" + this.id + "_newName\">Name</label><input id=\"" + this.id + "_newName\" class=\"tx-Rd-input\" name=\"tplName\" required maxlength=\"255\" value=\"New report\" autofocus>"
                + "<fieldset class=\"tx-Rd-cards\"><legend class=\"tx-Rd-label\">Start from</legend>"
                + card( "list", "fa-table", "List", "A title and a table of all rows with totals.", true )
                + card( "grouped", "fa-indent", "Grouped list", "Rows in groups, each with a subtotal." )
                + card( "document", "fa-file-text-o", "One page per record", "For example one invoice per order: fields, lines and a total." )
                + card( "blank", "fa-file-o", "Blank", "Only a title. Add blocks yourself." )
                + "</fieldset>"
                + "<div class=\"tx-Rd-dlgGroup\"><label class=\"tx-Rd-label\" for=\"" + this.id + "_newGroup\">Group or page by</label>"
                + "<select id=\"" + this.id + "_newGroup\" class=\"tx-Rd-input\" name=\"group\">" + options( this.colOptions(), guess ) + "</select></div>"
                + "<div class=\"tx-Rd-dlgButtons\"><button type=\"submit\" class=\"tx-Rd-btn\" value=\"cancel\" formnovalidate>Cancel</button>"
                + "<button type=\"submit\" class=\"tx-Rd-btn tx-Rd-btn--primary\" value=\"ok\">Create</button></div>",
                form => {
                    const kind = form.kind.value, name = form.tplName.value.trim() || "New report";
                    this.setTemplate( starter( kind, this.data, form.group.value ), { id: null, name, description: "", shared: false, mine: true } );
                    this.saved = null;   // a new template counts as unsaved
                    this.updateBar();
                }
            );
            // the column choice only matters for the grouped starters
            const form = this.dialog.querySelector( "form" );
            const sync = () => { form.querySelector( ".tx-Rd-dlgGroup" ).hidden = [ "grouped", "document" ].indexOf( form.kind.value ) < 0; };
            form.addEventListener( "change", sync );
            sync();
            } );
        }

        saveDialog( asNew ) {
            const m = this.meta, name = asNew || !m.mine ? ( m.name ? "Copy of " + m.name : "New report" ) : m.name;
            this.openDialog(
                "<h2 id=\"" + this.id + "_dlgTitle\">" + ( asNew ? "Save as a new template" : "Save template" ) + "</h2>"
                + "<label class=\"tx-Rd-label\" for=\"" + this.id + "_saveName\">Name</label><input id=\"" + this.id + "_saveName\" class=\"tx-Rd-input\" name=\"tplName\" required maxlength=\"255\" value=\"" + esc( name ) + "\" autofocus>"
                + "<label class=\"tx-Rd-label\" for=\"" + this.id + "_saveDesc\">Description</label><textarea id=\"" + this.id + "_saveDesc\" class=\"tx-Rd-input\" name=\"description\" rows=\"2\" maxlength=\"4000\">" + esc( m.description ) + "</textarea>"
                + ( this.cfg.sharing ? "<label class=\"tx-Rd-check\"><input type=\"checkbox\" name=\"shared\"" + ( m.shared && m.mine ? " checked" : "" ) + "><span>Share with everyone (others can use and copy it)</span></label>" : "" )
                + "<div class=\"tx-Rd-dlgButtons\"><button type=\"submit\" class=\"tx-Rd-btn\" value=\"cancel\" formnovalidate>Cancel</button>"
                + "<button type=\"submit\" class=\"tx-Rd-btn tx-Rd-btn--primary\" value=\"ok\">Save</button></div>",
                form => {
                    this.store( null, form.tplName.value.trim(), form.description.value, !!( form.shared && form.shared.checked ) );
                }
            );
        }

        save() {
            if ( !this.editable ) { return Promise.resolve(); }
            if ( !this.meta.id || !this.meta.mine ) { this.saveDialog( false ); return Promise.resolve(); }
            return this.store( this.meta.id, this.meta.name, this.meta.description, this.meta.shared );
        }

        store( id, name, description, shared ) {
            if ( !name ) { this.showError( "Enter a name for the template." ); return Promise.resolve(); }
            const btn = this.bar.querySelector( "[data-act=\"save\"]" );
            btn.disabled = true;
            this.setStatus( "Saving..." );
            return this.ajax( "SAVE", { x02: id || "", x03: name, x04: description || "", x05: shared ? "Y" : "N", p_clob_01: JSON.stringify( this.tpl ) } ).then( res => {
                this.templates = res.templates || this.templates;
                this.meta = { id: res.id, name, description: description || "", shared: !!shared, mine: true };
                this.snapshot( true );
                this.fillTemplateSelect();
                this.updateBar();
                this.renderProps();
                this.announce( "Saved" );
                apex.message.showPageSuccess( "Template \"" + name + "\" saved." );
                if ( this.cfg.templateItem ) { apex.item( this.cfg.templateItem ).setValue( String( res.id ) ); }
                apex.event.trigger( "#" + this.id, "tx-report-save", { id: res.id, name } );
            } ).catch( e => { this.setStatus( "" ); this.updateBar(); this.showError( e ); } ).finally( () => { btn.disabled = false; } );
        }

        remove() {
            const m = this.meta;
            if ( !m.id || !m.mine ) { return; }
            apex.message.confirm( "Delete the template \"" + m.name + "\"? This cannot be undone.", ok => {
                if ( !ok ) { return; }
                this.ajax( "DELETE", { x02: m.id } ).then( res => {
                    this.templates = res.templates || [];
                    this.saved = this.current;   // nothing left to keep
                    this.meta = { id: null, name: "", description: "", shared: false, mine: true };
                    apex.message.showPageSuccess( "Template \"" + m.name + "\" deleted." );
                    this.openInitial();
                } ).catch( e => this.showError( e ) );
            } );
        }

        // ---------- download ----------
        fileName( model, ext ) {
            const base = resolveText( this.tpl.settings.fileName || "{REPORT}", model.doc ).replace( /[\\/:*?"<>|\u0000-\u001f]+/g, " " ).replace( /\s+/g, " " ).trim();
            return ( base || "report" ).slice( 0, 120 ) + "." + ext;
        }

        download( format, btn ) {
            const info = FORMAT_INFO[ format ];
            if ( !info ) { return Promise.reject( new Error( "Unknown format " + format ) ); }
            btn = btn || this.bar.querySelector( "[data-format=\"" + format + "\"]" );
            if ( btn ) { btn.disabled = true; btn.classList.add( "is-busy" ); }
            const model = build( this.tpl, this.data, { name: this.meta.name || "Report", user: this.cfg.user } );
            const name = this.fileName( model, info.ext );
            return loadLib( this.cfg.filePrefix, format )
                .then( lib => format === "pdf" ? toPdf( model, lib ) : format === "docx" ? toDocx( model, lib ) : toXlsx( model, lib ) )
                .then( blob => {
                    saveBlob( blob, name );
                    apex.event.trigger( "#" + this.id, "tx-report-generate", { format, fileName: name, templateId: this.meta.id, templateName: this.meta.name } );
                    return blob;
                } )
                .catch( e => { apex.debug.error( "Report Designer", e ); this.showError( "The " + info.label + " file could not be created: " + ( e && e.message || e ) ); } )
                .finally( () => { if ( btn ) { btn.disabled = false; btn.classList.remove( "is-busy" ); } } );
        }

        // ---------- region interface ----------
        registerRegion() {
            apex.region.create( this.id, {
                type: "TechRevXReportDesigner",
                refresh: () => this.loadData(),
                generate: format => this.download( format ),
                openTemplate: id => this.openTemplate( id ),
                getTemplate: () => clone( this.tpl ),
                setTemplate: json => { this.setTemplate( typeof json === "string" ? JSON.parse( json ) : json ); this.saved = null; this.updateBar(); },
                save: () => this.save(),
                focus: () => { const f = this.el.querySelector( ".tx-Rd-tplSelect" ); if ( f ) { f.focus(); } }
            } );
        }
    }

    ns.reportDesigner = {
        init: regionId => new Designer( regionId ),
        // for tests and other code: the model and the renderers without the designer
        normalize, starter, build, toPdf, toDocx, toXlsx
    };

} )( apex.jQuery, apex );
