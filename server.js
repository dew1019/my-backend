require('dotenv').config();

const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const mongoose = require('mongoose');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const jwt = require('jsonwebtoken');

/* ------------------------- Basic config ------------------------- */
const app = express();
const PORT = process.env.PORT || 5003;
const isProd = process.env.NODE_ENV === 'production';

const PUBLIC_WEB_URL = (process.env.PUBLIC_WEB_URL || 'https://theglobalbpo.vercel.app').replace(/\/+$/, '');

/* ------------------------- Fail-fast in prod ------------------------- */
const requiredInProd = ['JWT_SECRET', 'MONGO_URI'];
const missing = requiredInProd.filter(k => !process.env[k]);
if (isProd && missing.length) {
    console.error(`❌ Missing required env in production: ${missing.join(', ')}`);
    process.exit(1);
}

/* ------------------------- Middleware ------------------------- */
const rawAllowed = (process.env.CORS_ORIGINS ||
    'http://localhost:3000,http://localhost:3001,https://theglobalbpo.vercel.app'
).split(',').map(s => s.trim()).filter(Boolean);

const allowedSet = new Set(rawAllowed);
function originAllowed(origin) {
    if (!origin) return true; // Postman/curl
    try {
        const u = new URL(origin);
        if (allowedSet.has(origin)) return true;
        // Allow Vercel previews like https://*.vercel.app
        if (u.hostname.endsWith('.vercel.app')) return true;
        if (u.hostname === 'www.theglobalbpo.com') return true;
        return false;
    } catch {
        return allowedSet.has(origin);
    }
}

app.use(cors({
    origin: (origin, cb) => cb(null, originAllowed(origin)),
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    credentials: true,
    optionsSuccessStatus: 204
}));

app.use(bodyParser.json({ limit: '25mb' }));

/* ------------------------- Storage ------------------------- */
const ROOT_DIR = __dirname;
const PDF_DIR = path.join(ROOT_DIR, 'pdfs');
const SIG_DIR = path.join(ROOT_DIR, 'signatures');
const PDF_TEMPLATES_DIR = path.join(ROOT_DIR, 'pdf-templates');
for (const d of [PDF_DIR, SIG_DIR, PDF_TEMPLATES_DIR]) if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });

/* ------------------------- Database ------------------------- */
const MONGO_URI = process.env.MONGO_URI || 'mongodb://localhost:27017/bpo_service_db';
mongoose.connect(MONGO_URI, { serverSelectionTimeoutMS: 15000 })
    .then(() => console.log('✅ Connected to MongoDB'))
    .catch(err => console.error('❌ MongoDB error:', err?.message || err));

/* ------------------------- Mailer (Prod: SendGrid, Dev: Gmail) ------------------------- */
let sendMailImpl;
if (process.env.SENDGRID_API_KEY) {
    const sg = require('@sendgrid/mail');
    sg.setApiKey(process.env.SENDGRID_API_KEY);
    const MAIL_FROM = process.env.MAIL_FROM || process.env.EMAIL_USER || 'no-reply@theglobalbpo.com';

    sendMailImpl = async ({ to, subject, text, html, attachments }) => {
        const files = (attachments || []).map(a => ({
            content: fs.readFileSync(a.path).toString('base64'),
            filename: a.filename,
            type: 'application/pdf',
            disposition: 'attachment'
        }));
        await sg.send({
            to,
            from: MAIL_FROM,
            subject,
            text,
            html,
            attachments: files
        });
    };
} else {
    // Local/dev Gmail via SMTP (App Password required). In prod this may be blocked by host firewalls.
    const user = process.env.EMAIL_USER;
    const pass = process.env.EMAIL_PASS;
    const gmailTransport = (user && pass)
        ? nodemailer.createTransport({
            service: 'gmail',
            auth: { user, pass },
            pool: true,
            maxConnections: 2,
            connectionTimeout: 12000,
            greetingTimeout: 12000,
            socketTimeout: 12000,
        })
        : nodemailer.createTransport({ jsonTransport: true });

    sendMailImpl = async (opts) => {
        await gmailTransport.sendMail({ from: process.env.EMAIL_USER, ...opts });
    };
}

async function safeSendMail(opts) {
    try {
        await sendMailImpl(opts);
    } catch (err) {
        const detail = err?.response?.body || err?.message || String(err);
        console.error('✉️  Mail send failed:', detail);
    }
}

/* ------------------------- Schemas ------------------------- */
const clientSchema = new mongoose.Schema({
    businessName: String,
    email: String,
    phone: String,
    createdAt: { type: Date, default: Date.now },
});

const directorSlotSchema = new mongoose.Schema({
    label: String,
    email: String,
    directorSignToken: String,
    signToken: String,
    signatureImagePath: String,
    signature: String,
    signed: { type: Boolean, default: false },
    signedDate: Date,
    signedPdfPath: String,
}, { _id: false });

const agreementDocumentSchema = new mongoose.Schema({
    name: String,
    draftPdfPath: String,
    clientSignedPdfPath: String,
    finalSignedPdfPath: String,
    clientSignatureImagePath: String,

    directorSignatureImagePath: String,
    directorSignToken: String,
    directorSigned: { type: Boolean, default: false },
    directorSignedDate: Date,
    directorSignature: String,

    clientSignToken: String,
    clientSigned: { type: Boolean, default: false },
    clientSignedDate: Date,
    clientSignature: String,

    directors: [directorSlotSchema],
}, { _id: false });

const agreementSchema = new mongoose.Schema({
    businessName: String,
    tradingName: String,
    clientFullName: String,
    dateOfBirth: String,
    email: String,
    phone: String,

    registeredOffice: String,
    registeredPostCode: String,
    postalAddress: String,
    postalPostCode: String,
    businessAddress: String,
    businessPostCode: String,

    departmentContact: String,
    contactNumber: String,
    departmentEmail: String,
    officeNumber: String,

    nameOfDirector: String,
    addressOfDirector: String,
    driversLicense: String,

    acn_abn: String,
    mainService: String,
    contractStartDate: String,
    JobType: String,

    businessType: String,
    ACN: String,
    ABN: String,
    address: String,
    city: String,
    state: String,
    postalCode: String,
    website: String,
    additionalNotes: String,

    submittedAt: { type: Date, default: Date.now },

    documents: [agreementDocumentSchema],

    clientSignToken: String,
    directorSignToken: String,
    clientSigned: { type: Boolean, default: false },
    directorSigned: { type: Boolean, default: false },
    clientSignedDate: Date,
    directorSignedDate: Date,
    clientSignature: String,
    directorSignature: String,
});

const summarySchema = new mongoose.Schema({
    businessName: String,
    email: String,
    phone: String,
    service: String,
    plan: { name: String, price: Number },
    addOns: [{ name: String, price: Number }],
    total: Number,
    createdAt: { type: Date, default: Date.now },
});

const Client = mongoose.model('Client', clientSchema, 'clients');
const Agreement = mongoose.model('Agreement', agreementSchema, 'agreements');
const Summary = mongoose.model('Summary', summarySchema, 'summaries');

/* ------------------------- Utils ------------------------- */
function dataUrlToBuffer(dataUrl) {
    if (!dataUrl) return null;
    const base64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl;
    return Buffer.from(base64, 'base64');
}
const nowIso = () => new Date().toISOString();
function nowLocal() {
    try { return new Date().toLocaleString('en-AU', { timeZone: process.env.TIMEZONE || 'Australia/Melbourne' }); }
    catch { return new Date().toLocaleString(); }
}
function safeFolderName(name = 'Client') {
    let s = (name || 'Client')
        .normalize('NFKD')
        .replace(/[^\w\-]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    s = s.replace(/[.#]+$/g, '').replace(/^\.+/g, '');
    if (!s || s.toLowerCase() === 'forms') s = 'Client';
    if (s.length > 100) s = s.slice(0, 100).trim();
    return s;
}

/* ------------------------- PDF helpers ------------------------- */
function getSafePage(pdfDoc, wantedPage, label = 'page') {
    const count = pdfDoc.getPageCount();
    let idx;
    if (wantedPage === 'last') idx = count - 1;
    else {
        const zero = Math.max(0, (wantedPage || 1) - 1);
        idx = Math.min(zero, count - 1);
    }
    if (count && typeof wantedPage === 'number' && idx !== wantedPage - 1) {
        console.warn(`[PDF] ${label}: requested p${wantedPage}, but PDF has ${count} page(s). Using p${idx + 1}.`);
    }
    return pdfDoc.getPage(idx);
}

function resolveSignatureCoords(pdfDoc, templateCoords, override, label) {
    const merged = { ...(templateCoords || {}), ...(override || {}) };
    const page = merged.page != null ? merged.page : templateCoords?.page ?? 'last';
    const p = getSafePage(pdfDoc, page, label);
    const width = Math.max(1, merged.width || 150);
    const height = Math.max(1, merged.height || 50);
    const origin = merged.origin || 'top-left';
    let x = Number.isFinite(merged.x) ? merged.x : 0;
    let y = Number.isFinite(merged.y) ? merged.y : 0;
    x = Math.min(Math.max(0, x), p.getWidth() - width);
    y = Math.min(Math.max(0, y), p.getHeight() - height);
    return { page, x, y, width, height, origin };
}

async function drawFieldsOnPdf(templatePath, fieldMap, inputData, outPath) {
    const existingPdfBytes = fs.readFileSync(templatePath);
    const pdfDoc = await PDFDocument.load(existingPdfBytes);
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    function wrapLines(text, size, maxWidth) {
        if (!maxWidth) return [text];
        const words = String(text).split(/\s+/);
        const lines = [];
        let line = '';
        for (const w of words) {
            const test = line ? line + ' ' + w : w;
            if (font.widthOfTextAtSize(test, size) <= maxWidth) line = test;
            else { if (line) lines.push(line); line = w; }
        }
        if (line) lines.push(line);
        return lines;
    }

    function drawOnePlacement(p, spec, value) {
        const { size = 10, color = rgb(0, 0, 0), maxWidth, lineHeight = 1.2 } = spec;
        let { x, y } = spec;
        if (spec.origin === 'top-left') y = p.getHeight() - spec.y;
        if (maxWidth) {
            const lines = wrapLines(value, size, maxWidth);
            lines.forEach((line, i) => p.drawText(line, { x, y: y - i * (size * lineHeight), size, font, color }));
        } else {
            p.drawText(value, { x, y, size, font, color });
        }
    }

    for (const [key, rawSpec] of Object.entries(fieldMap)) {
        const value = (inputData?.[key] ?? '').toString();
        if (!value) continue;
        const placements = Array.isArray(rawSpec) ? rawSpec : [rawSpec];
        for (const spec of placements) {
            const p = getSafePage(pdfDoc, spec.page || 1, `field "${key}"`);
            drawOnePlacement(p, spec, value);
        }
    }
    fs.writeFileSync(outPath, await pdfDoc.save());
    return outPath;
}

async function stampSignatureWithTime(inPath, signatureDataUrl, templateCoords, overrideCoords, outPath, label = 'Signed at') {
    const pdfDoc = await PDFDocument.load(fs.readFileSync(inPath));
    const coords = resolveSignatureCoords(pdfDoc, templateCoords, overrideCoords, label);
    const p = getSafePage(pdfDoc, coords.page, label);

    const png = await pdfDoc.embedPng(dataUrlToBuffer(signatureDataUrl));
    const width = coords.width || png.width;
    const height = coords.height || png.height;

    let { x, y } = coords;
    if (coords.origin === 'top-left') y = p.getHeight() - y - height;

    p.drawImage(png, { x, y, width, height });
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    p.drawText(`${label}: ${nowLocal()}`, { x, y: y - 14, size: 10, font, color: rgb(0, 0, 0) });

    fs.writeFileSync(outPath, await pdfDoc.save());
    return outPath;
}

/* ------------------------- Templates ------------------------- */
const TEMPLATES = [
    {
        name: 'ServiceAgreement',
        path: 'Engagement-letter.pdf',
        fieldMap: {
            businessName: { page: 2, x: 171, y: 563, size: 10, origin: 'top-left' },
            tradingName: { page: 2, x: 171, y: 540, size: 10, origin: 'top-left' },
            clientFullName: { page: 2, x: 171, y: 484, size: 10, origin: 'top-left' },
            dateOfBirth: { page: 2, x: 171, y: 462, size: 10, origin: 'top-left' },
            email: { page: 2, x: 169, y: 450, size: 10, origin: 'top-left', maxWidth: 260 },
            phone: { page: 2, x: 405, y: 484, size: 10, origin: 'top-left' },

            registeredOffice: { page: 2, x: 171, y: 395, size: 10, origin: 'top-left' },
            registeredPostCode: { page: 2, x: 405, y: 395, size: 10, origin: 'top-left' },
            postalAddress: { page: 2, x: 171, y: 347, size: 10, origin: 'top-left' },
            postalPostCode: { page: 2, x: 405, y: 347, size: 10, origin: 'top-left' },
            businessAddress: { page: 2, x: 171, y: 297, size: 10, origin: 'top-left' },
            businessPostCode: { page: 2, x: 405, y: 297, size: 10, origin: 'top-left' },

            departmentContact: { page: 2, x: 171, y: 259, size: 10, origin: 'top-left' },
            contactNumber: { page: 2, x: 405, y: 259, size: 10, origin: 'top-left' },
            departmentEmail: { page: 2, x: 171, y: 237, size: 10, origin: 'top-left' },
            officeNumber: { page: 2, x: 405, y: 450, size: 10, origin: 'top-left' },

            nameOfDirector: { page: 2, x: 171, y: 217, size: 10, origin: 'top-left' },
            addressOfDirector: { page: 2, x: 405, y: 217, size: 10, origin: 'top-left' },
            driversLicense: { page: 2, x: 171, y: 195, size: 10, origin: 'top-left' },

            ACN: { page: 2, x: 405, y: 566, size: 10, origin: 'top-left' },
            ABN: { page: 2, x: 405, y: 521, size: 10, origin: 'top-left' },
            acn_abn: { page: 2, x: 171, y: 170, size: 10, origin: 'top-left' },

            mainService: { page: 4, x: 319, y: 554, size: 11, origin: 'top-left' },
            contractStartDate: { page: 3, x: 182, y: 428, size: 10, origin: 'top-left' },
            JobType: { page: 3, x: 182, y: 400, size: 10, origin: 'top-left' },

            businessType: { page: 2, x: 171, y: 150, size: 10, origin: 'top-left' },
            address: { page: 2, x: 171, y: 130, size: 10, origin: 'top-left' },
            city: { page: 2, x: 171, y: 110, size: 10, origin: 'top-left' },
            state: { page: 2, x: 171, y: 90, size: 10, origin: 'top-left' },
            postalCode: { page: 2, x: 405, y: 90, size: 10, origin: 'top-left' },
            website: { page: 2, x: 171, y: 70, size: 10, origin: 'top-left' },
            additionalNotes: { page: 2, x: 171, y: 50, size: 10, origin: 'top-left', maxWidth: 260 },

            // Summary on first PDF
            planName: { page: 1, x: 150, y: 620, size: 10, origin: 'top-left' },
            planPrice: { page: 1, x: 420, y: 620, size: 10, origin: 'top-left' },
            addOns: { page: 1, x: 150, y: 600, size: 10, origin: 'top-left', maxWidth: 360 },
            total: { page: 1, x: 420, y: 580, size: 12, origin: 'top-left' },
        },
        clientSig: { page: 6, x: 338, y: 125, width: 150, height: 50, origin: 'top-left' },
        directorSig: { page: 6, x: 338, y: 72, width: 150, height: 50, origin: 'top-left' },
    },

    {
        name: 'CustomerInformation',
        path: 'guarantee-page.pdf',
        fieldMap: {
            businessName: { page: 1, x: 388, y: 616, size: 10, origin: 'top-left' },
            ACN: { page: 1, x: 101, y: 605, size: 10, origin: 'top-left' },
            ABN: { page: 1, x: 101, y: 563, size: 10, origin: 'top-left' },
            directorName: { page: 1, x: 384, y: 102, size: 10, origin: 'top-left' },
            submittedAt: { page: 1, x: 338, y: 74, size: 10, origin: 'top-left' },
        },
        clientSig: { page: 6, x: 338, y: 125, width: 150, height: 50, origin: 'top-left' },
        directorSig: { page: 6, x: 64, y: 466, width: 150, height: 50, origin: 'top-left' },
        director1Sig: { page: 4, x: 350, y: 466, width: 150, height: 50, origin: 'top-left' },
    },

    {
        name: 'PricingSchedule',
        path: 'privacy-policy.pdf',
        fieldMap: {
            businessName: { page: 1, x: 150, y: 660, size: 10, origin: 'top-left' },
            service: { page: 1, x: 150, y: 640, size: 10, origin: 'top-left' },
            planName: { page: 1, x: 150, y: 620, size: 10, origin: 'top-left' },
            planPrice: { page: 1, x: 420, y: 620, size: 10, origin: 'top-left' },
            addOnsSummary: { page: 1, x: 150, y: 600, size: 10, origin: 'top-left', maxWidth: 360 },
            total: { page: 1, x: 420, y: 580, size: 12, origin: 'top-left' },
        },
        clientSig: { page: 6, x: 420, y: 90, width: 150, height: 50, origin: 'top-left' },
        directorSig: { page: 6, x: 420, y: 40, width: 150, height: 50, origin: 'top-left' },
    },

    {
        name: 'ConfidentialityAgreement',
        path: 'terms.pdf',
        fieldMap: {
            businessName: { page: 4, x: 120, y: 362, size: 10, origin: 'top-left' },
            date: { page: 4, x: 430, y: 630, size: 10, origin: 'top-left' },
            ACN: { page: 4, x: 406, y: 608, size: 10, origin: 'top-left' },
            ABN: { page: 4, x: 406, y: 563, size: 10, origin: 'top-left' },
        },
        clientSig: { page: 6, x: 380, y: 80, width: 150, height: 50, origin: 'top-left' },
        directorSig: { page: 4, x: 63, y: 466, width: 150, height: 50, origin: 'top-left' },
        director1Sig: { page: 4, x: 350, y: 466, width: 150, height: 50, origin: 'top-left' },
    },
];

/* ------------------------- Template existence checks ------------------------- */
const TEMPLATE_FILES = TEMPLATES.map(t => t.path);
function listTemplatesStatus() {
    return TEMPLATE_FILES.map(name => {
        const p = path.join(PDF_TEMPLATES_DIR, name);
        return { name, path: p, exists: fs.existsSync(p) };
    });
}
(function checkTemplatesAtBoot() {
    const status = listTemplatesStatus();
    const missing = status.filter(s => !s.exists);
    if (missing.length) {
        console.error('❌ Missing PDF templates:', missing.map(m => m.name).join(', '));
        console.error(`Place them under ${PDF_TEMPLATES_DIR} (case-sensitive) and redeploy.`);
    } else {
        console.log('✅ All PDF templates found.');
    }
})();

/* ------------------------- Self-heal drafts ------------------------- */
async function ensureDraftExists(agreement, doc) {
    const template = TEMPLATES.find(t => t.name === doc.name);
    if (!template) throw new Error(`Template not found for ${doc.name}`);

    const templateStatus = listTemplatesStatus().find(s => s.name === template.path);
    if (!templateStatus?.exists) throw new Error(`Template file missing: ${template.path}`);

    const templatePath = path.join(PDF_TEMPLATES_DIR, template.path);
    const summary = await Summary.findOne({ email: agreement.email }).sort({ createdAt: -1 });
    const input = buildTemplateInput(doc.name, agreement.toObject(), summary);

    const outDraft = path.join(PDF_DIR, `rehydrated-${doc.name}-${Date.now()}.pdf`);
    await drawFieldsOnPdf(templatePath, template.fieldMap, input, outDraft);

    if (!doc.draftPdfPath) doc.draftPdfPath = outDraft;
    await agreement.save();
    return outDraft;
}

/* ------------------------- Template input mapping ------------------------- */
function buildTemplateInput(templateName, agreementData, summaryData) {
    const base = {
        businessName: agreementData.businessName || '',
        tradingName: agreementData.tradingName || '',
        clientFullName: agreementData.clientFullName || '',
        dateOfBirth: agreementData.dateOfBirth || '',
        email: agreementData.email || '',
        phone: agreementData.phone || '',
        registeredOffice: agreementData.registeredOffice || '',
        registeredPostCode: agreementData.registeredPostCode || '',
        postalAddress: agreementData.postalAddress || '',
        postalPostCode: agreementData.postalPostCode || '',
        businessAddress: agreementData.businessAddress || '',
        businessPostCode: agreementData.businessPostCode || '',
        departmentContact: agreementData.departmentContact || '',
        contactNumber: agreementData.contactNumber || '',
        departmentEmail: agreementData.departmentEmail || '',
        officeNumber: agreementData.officeNumber || '',
        nameOfDirector: agreementData.nameOfDirector || '',
        addressOfDirector: agreementData.addressOfDirector || '',
        driversLicense: agreementData.driversLicense || '',
        acn_abn: agreementData.acn_abn || '',
        mainService: agreementData.mainService || '',
        contractStart: agreementData.contractStartDate || '',
        contractStartDate: agreementData.contractStartDate || '',

        businessType: agreementData.businessType || '',
        ACN: agreementData.ACN || '',
        ABN: agreementData.ABN || '',
        address: agreementData.address || '',
        city: agreementData.city || '',
        state: agreementData.state || '',
        postalCode: agreementData.postalCode || '',
        website: agreementData.website || '',
        additionalNotes: agreementData.additionalNotes || '',

        submittedAt: new Date(agreementData.submittedAt || Date.now()).toLocaleString(),
    };

    base.companyName = base.businessName;
    base.dearName = base.businessName;
    base.customerName = base.businessName;
    base.clientEmail = base.email;
    base.mobileNumber = base.phone;
    base.postCode = base.registeredPostCode || base.postalPostCode || base.businessPostCode || '';
    base.p4CompanyName = base.businessName;
    base.servicesAssist = base.mainService || '';
    base.directorName = base.nameOfDirector;
    base.p5ClientName = base.clientFullName;
    base.p5OtherName = base.directorName;

    const dateOnly = base.submittedAt.split(',')[0];
    base.p5Date1 = dateOnly;
    base.p5Date2 = dateOnly;

    base.guarantorCompany = base.businessName;
    base.guarantorACNABN = base.ACN || base.ABN || base.acn_abn;

    if (templateName === 'PricingSchedule' && summaryData) {
        const addOnsSummary = (summaryData.addOns || []).map(a => `${a.name} ($${a.price})`).join(', ');
        base.service = summaryData.service || base.servicesAssist || '';
        base.planName = summaryData.plan?.name || '';
        base.planPrice = typeof summaryData.plan?.price === 'number' ? `$${summaryData.plan.price.toFixed(2)}` : '';
        base.addOnsSummary = addOnsSummary;
        base.total = typeof summaryData.total === 'number' ? `$${summaryData.total.toFixed(2)}` : '';
    }

    if (templateName === 'ConfidentialityAgreement') base.executedDate = new Date().toLocaleDateString();

    return base;
}

function getDirectorSigSpec(template, directorIndex) {
    if (!template) return null;
    if (directorIndex === 0 && template.directorSig) return template.directorSig;
    if (directorIndex === 1) {
        if (template.director1Sig) return template.director1Sig;
        if (template.directorSig) {
            const { page, x, y, width, height, origin } = template.directorSig;
            return { page, x, y: y + 60, width, height, origin };
        }
    }
    return template.directorSig || template.director1Sig || null;
}

/* ------------------------- Microsoft Graph / SharePoint ------------------------- */
const GRAPH_TENANT = (process.env.GRAPH_TENANT_ID || '').trim();
const GRAPH_CLIENT = (process.env.GRAPH_CLIENT_ID || '').trim();
const GRAPH_SECRET = (process.env.GRAPH_CLIENT_SECRET || '').trim();
const SITE_ID = (process.env.GRAPH_SITE_ID || '').trim();
const DRIVE_ID = (process.env.GRAPH_DRIVE_ID || '').trim();
const SP_BASE_PATH = (process.env.SP_BASE_PATH || 'Agreements').replace(/^\/|\/$/g, '');

app.get('/api/debug/graph-env', (req, res) => {
    res.json({
        GRAPH_TENANT_ID: Boolean(process.env.GRAPH_TENANT_ID),
        GRAPH_CLIENT_ID: Boolean(process.env.GRAPH_CLIENT_ID),
        GRAPH_CLIENT_SECRET: Boolean(process.env.GRAPH_CLIENT_SECRET && process.env.GRAPH_CLIENT_SECRET.length > 5),
        GRAPH_SITE_ID: Boolean(process.env.GRAPH_SITE_ID),
        GRAPH_DRIVE_ID: Boolean(process.env.GRAPH_DRIVE_ID),
        SP_BASE_PATH: process.env.SP_BASE_PATH || 'Agreements'
    });
});

async function getAppToken() {
    const url = `https://login.microsoftonline.com/${GRAPH_TENANT}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('client_id', GRAPH_CLIENT);
    params.append('client_secret', GRAPH_SECRET);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('grant_type', 'client_credentials');
    const { data } = await axios.post(url, params);
    return data.access_token;
}

async function graphRequest(token, method, url, data, headers = {}) {
    return axios({
        method,
        url: `https://graph.microsoft.com/v1.0${url}`,
        headers: { Authorization: `Bearer ${token}`, ...headers },
        data
    });
}

async function ensureFolderPath(token, driveId, segments) {
    let currentPath = '';
    for (const seg of segments) {
        currentPath = currentPath ? `${currentPath}/${seg}` : seg;
        try {
            await graphRequest(token, 'GET', `/drives/${driveId}/root:/${encodeURIComponent(currentPath)}`);
        } catch (e) {
            if (e?.response?.status !== 404) throw e;
            const parentPath = currentPath.split('/').slice(0, -1).join('/');
            const parentUrl = parentPath
                ? `/drives/${driveId}/root:/${encodeURIComponent(parentPath)}:/children`
                : `/drives/${driveId}/root/children`;
            await graphRequest(token, 'POST', parentUrl, {
                name: seg, folder: {}, "@microsoft.graph.conflictBehavior": "replace"
            });
        }
    }
    return currentPath;
}

async function uploadSmall(token, driveId, fullPath, fileBuffer) {
    const url = `/drives/${driveId}/root:/${encodeURIComponent(fullPath)}:/content`;
    await graphRequest(token, 'PUT', url, fileBuffer, { 'Content-Type': 'application/octet-stream' });
}
async function uploadLarge(token, driveId, fullPath, filePath) {
    const session = await graphRequest(
        token, 'POST',
        `/drives/${driveId}/root:/${encodeURIComponent(fullPath)}:/createUploadSession`,
        { item: { "@microsoft.graph.conflictBehavior": "replace", name: path.basename(fullPath) } }
    );
    const uploadUrl = session.data.uploadUrl;
    const CHUNK = 5 * 1024 * 1024;
    const size = fs.statSync(filePath).size;
    const fd = fs.openSync(filePath, 'r');

    let next = 0;
    while (next < size) {
        const left = size - next;
        const toRead = Math.min(CHUNK, left);
        const buffer = Buffer.alloc(toRead);
        fs.readSync(fd, buffer, 0, toRead, next);
        const from = next;
        const to = next + toRead - 1;

        await axios.put(uploadUrl, buffer, {
            headers: {
                'Content-Length': toRead,
                'Content-Range': `bytes ${from}-${to}/${size}`
            }
        });
        next += toRead;
    }
    fs.closeSync(fd);
}

async function uploadFile(token, driveId, folderPath, localFilePath) {
    if (!fs.existsSync(localFilePath)) return;
    const name = path.basename(localFilePath);
    const fullPath = folderPath ? `${folderPath}/${name}` : name;
    const size = fs.statSync(localFilePath).size;

    if (size <= 4 * 1024 * 1024) {
        const buffer = fs.readFileSync(localFilePath);
        await uploadSmall(token, driveId, fullPath, buffer);
    } else {
        await uploadLarge(token, driveId, fullPath, localFilePath);
    }
}

async function uploadAgreementToSharePoint(agreement) {
    if (!GRAPH_TENANT || !GRAPH_CLIENT || !GRAPH_SECRET || !SITE_ID || !DRIVE_ID) {
        console.warn('⚠️ Graph env not set; skipping SharePoint upload.');
        return;
    }
    const token = await getAppToken();

    const clientFolder = safeFolderName(agreement.businessName || 'Client');
    const segments = [SP_BASE_PATH, clientFolder].filter(Boolean);
    const folderPath = await ensureFolderPath(token, DRIVE_ID, segments);

    for (const d of agreement.documents || []) {
        for (const p of [d.finalSignedPdfPath, d.clientSignedPdfPath, d.draftPdfPath]) {
            if (p && fs.existsSync(p)) await uploadFile(token, DRIVE_ID, folderPath, p);
        }
    }
    for (const d of agreement.documents || []) {
        if (d.clientSignatureImagePath && fs.existsSync(d.clientSignatureImagePath)) {
            await uploadFile(token, DRIVE_ID, folderPath, d.clientSignatureImagePath);
        }
        for (const dir of (d.directors || [])) {
            if (dir.signatureImagePath && fs.existsSync(dir.signatureImagePath)) {
                await uploadFile(token, DRIVE_ID, folderPath, dir.signatureImagePath);
            }
        }
    }
    console.log(`✅ Uploaded PDFs & signatures to SharePoint: /${folderPath}`);
}

/* ------------------------- helpers ------------------------- */
function getDirectorEmails() {
    const multi = (process.env.DIRECTOR_EMAILS || '')
        .split(',').map(s => s.trim()).filter(Boolean);
    if (multi.length) return multi;
    return (process.env.DIRECTOR_EMAIL ? [process.env.DIRECTOR_EMAIL] : []);
}

/* ------------------------- routes ------------------------- */

// New client login
app.post('/api/new-client-login', async (req, res) => {
    const { businessName, email, phone } = req.body;
    try {
        await new Client({ businessName, email, phone }).save();

        await safeSendMail({
            to: process.env.AGREEMENTS_INBOX || process.env.EMAIL_USER,
            subject: '🚀 New Client Logged In',
            text: `Business: ${businessName}\nEmail: ${email}\nPhone: ${phone}`,
        });

        res.status(200).json({ message: 'Client saved and emails sent.' });
    } catch (err) {
        console.error('❌ Error saving client:', err);
        res.status(500).json({ message: 'Error occurred while saving client.' });
    }
});

// Tiny helper for the signing UI header
app.get('/api/sign/session/:token', async (req, res) => {
    try {
        const decoded = jwt.verify(req.params.token, process.env.JWT_SECRET);
        const ag = await Agreement.findById(decoded.agid).select('businessName email');
        if (!ag) return res.status(404).json({});
        res.json({ name: ag.businessName || '', email: ag.email || '' });
    } catch { res.json({}); }
});

// Save pricing summary
app.post('/api/save-pricing-summary', async (req, res) => {
    try {
        const { businessName, email, phone, service, plan, addOns, total } = req.body;
        if (!plan || !plan.name || typeof plan.price !== 'number')
            return res.status(400).json({ message: 'Invalid or missing plan data.' });
        await new Summary({ businessName, email, phone, service, plan, addOns, total }).save();
        res.status(200).json({ message: 'Pricing summary saved' });
    } catch (e) {
        console.error('❌ Error saving summary:', e);
        res.status(500).json({ message: 'Failed to save summary' });
    }
});

// Prefill the full pack and email client links
app.post('/api/submit-agreement', async (req, res) => {
    const b = req.body;
    const businessName = b.CompanyName ?? b.businessName;
    const clientFullName = b.ClientFullName ?? b.clientFullName;
    const email = b.ClientEmail ?? b.email;
    const phone = b.MobileNumber ?? b.phone;

    if (![businessName, clientFullName, email, phone].every(v => v && String(v).trim()))
        return res.status(400).json({ message: 'Missing required fields (businessName, clientFullName, email, phone).' });

    try {
        const idDigits = String(b.ACN_ABN || '').replace(/\D/g, '');
        const ACN = idDigits.length === 9 ? idDigits : (b.ACN || '');
        const ABN = idDigits.length === 11 ? idDigits : (b.ABN || '');

        const payload = {
            businessName,
            tradingName: b.TradingName ?? b.tradingName ?? '',
            clientFullName,
            dateOfBirth: b.DateOfBirth ?? b.dateOfBirth ?? '',
            email,
            phone,

            registeredOffice: b.RegisteredOffice ?? b.registeredOffice ?? '',
            registeredPostCode: b.RegisteredPostCode ?? b.registeredPostCode ?? '',
            postalAddress: b.PostalAddress ?? b.postalAddress ?? '',
            postalPostCode: b.PostalPostCode ?? b.postalPostCode ?? '',
            businessAddress: b.BusinessAddress ?? b.businessAddress ?? '',
            businessPostCode: b.BusinessPostCode ?? b.businessPostCode ?? '',

            departmentContact: b.DepartmentContact ?? b.departmentContact ?? '',
            contactNumber: b.ContactNumber ?? b.contactNumber ?? '',
            departmentEmail: b.DepartmentEmail ?? b.departmentEmail ?? '',
            officeNumber: b.OfficeNumber ?? b.officeNumber ?? '',

            nameOfDirector: b.NameOfDirector ?? b.nameOfDirector ?? '',
            addressOfDirector: b.AddressOfDirector ?? b.addressOfDirector ?? '',
            driversLicense: b.DriversLicense ?? b.driversLicense ?? '',

            acn_abn: b.ACN_ABN ?? `${b.ACN || ''}${b.ABN ? ' / ' + b.ABN : ''}`,
            ACN, ABN,

            mainService: b.MainService ?? b.mainService ?? '',
            contractStartDate: b.ContractStartDate ?? b.contractStartDate ?? '',
            JobType: b.JobType ?? b.jobType ?? '',

            businessType: b.businessType ?? '',
            address: b.address ?? '',
            city: b.city ?? '',
            state: b.state ?? '',
            postalCode: b.postalCode ?? '',
            website: b.website ?? '',
            additionalNotes: b.additionalNotes ?? '',

            submittedAt: new Date(),
        };

        const newAgreement = new Agreement(payload);
        const summary = await Summary.findOne({ email: payload.email }).sort({ createdAt: -1 });

        const documents = [];
        const signLinks = [];
        for (const t of TEMPLATES) {
            const outDraft = path.join(PDF_DIR, `draft-${t.name}-${Date.now()}.pdf`);
            const templateStatus = listTemplatesStatus().find(s => s.name === t.path);
            if (!templateStatus?.exists) throw new Error(`Template file missing: ${t.path}`);

            const templatePath = path.join(PDF_TEMPLATES_DIR, t.path);
            const input = buildTemplateInput(t.name, payload, summary);
            await drawFieldsOnPdf(templatePath, t.fieldMap, input, outDraft);

            const clientSignToken = jwt.sign(
                { agid: newAgreement._id.toString(), docName: t.name, role: 'client' },
                process.env.JWT_SECRET, { expiresIn: '7d' }
            );

            documents.push({
                name: t.name,
                draftPdfPath: outDraft,
                clientSigned: false,
                directorSigned: false,
                clientSignToken,
                directors: [],
            });

            signLinks.push(`• ${t.name}: ${PUBLIC_WEB_URL}/sign/${clientSignToken}`);
        }

        newAgreement.documents = documents;
        newAgreement.clientSignToken = documents[0].clientSignToken;
        await newAgreement.save();

        const attachments = documents.map(d => ({ filename: `${d.name}-Draft.pdf`, path: d.draftPdfPath }));
        await safeSendMail({
            to: newAgreement.email,
            bcc: process.env.AGREEMENTS_INBOX,
            subject: '📄 Your Prefilled Agreement Pack',
            text:
                `Hi ${newAgreement.businessName},\n\n` +
                `Your agreement pack is ready. Please sign each document using the links below:\n\n` +
                `${signLinks.join('\n')}\n\nKind regards,\nThe Global BPO Team`,
            attachments,
        });

        res.status(200).json({ message: 'Agreement saved, drafts generated.', signLinks });
    } catch (e) {
        console.error('❌ Error saving agreement:', e);
        res.status(500).json({ message: 'Failed to save and send agreement.' });
    }
});

// Self-healing PDF preview
app.get('/api/sign/preview/:token', async (req, res) => {
    try {
        const decoded = jwt.verify(req.params.token, process.env.JWT_SECRET);
        const ag = await Agreement.findById(decoded.agid);
        if (!ag) return res.status(404).send('Agreement not found');

        const doc = decoded.docName ? ag.documents.find(d => d.name === decoded.docName) : ag.documents[0];
        if (!doc) return res.status(404).send('Document not found');

        let filePath = (decoded.role === 'client')
            ? (doc.clientSignedPdfPath || doc.draftPdfPath || doc.finalSignedPdfPath)
            : (doc.finalSignedPdfPath || doc.clientSignedPdfPath || doc.draftPdfPath);

        if (!filePath || !fs.existsSync(filePath)) {
            filePath = await ensureDraftExists(ag, doc);
        }

        res.setHeader('Content-Type', 'application/pdf');
        fs.createReadStream(filePath).pipe(res);
    } catch (e) {
        console.error('preview error:', e);
        res.status(404).send('PDF not found');
    }
});

// Client signs a document
app.post('/api/sign/:token', async (req, res) => {
    const { signature, coords } = req.body;
    if (!signature) return res.status(400).json({ message: 'Missing signature' });
    try {
        const decoded = jwt.verify(req.params.token, process.env.JWT_SECRET);
        if (decoded.role !== 'client') return res.status(403).json({ message: 'Wrong link' });

        const ag = await Agreement.findById(decoded.agid);
        if (!ag) return res.status(404).json({ message: 'Agreement not found' });

        const doc = ag.documents.find(d => d.name === decoded.docName);
        if (!doc) return res.status(404).json({ message: 'Document not found' });
        if (doc.clientSigned) return res.status(400).json({ message: 'Document already signed' });

        const template = TEMPLATES.find(t => t.name === doc.name);
        if (!template) return res.status(500).json({ message: `Template not found for ${doc.name}` });

        let inFile = doc.draftPdfPath;
        if (!inFile || !fs.existsSync(inFile)) {
            inFile = await ensureDraftExists(ag, doc);
            doc.draftPdfPath = inFile;
            await ag.save();
        }

        const ts = nowIso().replace(/[:.]/g, '-');
        const safeBusiness = safeFolderName(ag.businessName || 'Client');

        const sigPath = path.join(SIG_DIR, `client-sig-${doc.name}-${safeBusiness}-${ts}-${ag._id}.png`);
        fs.writeFileSync(sigPath, dataUrlToBuffer(signature));

        const outFile = path.join(PDF_DIR, `client-${doc.name}-${ag._id}-${ts}.pdf`);
        await stampSignatureWithTime(inFile, signature, template.clientSig, coords, outFile, 'Client signed at');

        doc.clientSignatureImagePath = sigPath;
        doc.clientSignedPdfPath = outFile;
        doc.clientSigned = true;
        doc.clientSignature = signature;
        doc.clientSignedDate = new Date();
        await ag.save();

        const nextDoc = ag.documents.find(d => !d.clientSigned);
        if (nextDoc) {
            const nextToken = jwt.sign(
                { agid: ag._id.toString(), role: 'client', docName: nextDoc.name },
                process.env.JWT_SECRET, { expiresIn: '7d' }
            );

            nextDoc.clientSignToken = nextToken;
            await ag.save();
            return res.status(200).json({ nextDocToken: nextToken });
        }

        // All client docs done -> prepare director links & email
        const directorEmails = getDirectorEmails();
        for (const d of ag.documents) {
            d.directors = directorEmails.map((email, idx) => ({
                email,
                directorSignToken: jwt.sign(
                    { agid: ag._id.toString(), role: 'director', docName: d.name, directorIndex: idx },
                    process.env.JWT_SECRET, { expiresIn: '7d' }
                ),
                signed: false
            }));
        }
        ag.clientSigned = true;
        ag.clientSignedDate = new Date();
        await ag.save();

        for (let i = 0; i < directorEmails.length; i++) {
            const email = directorEmails[i];
            const directorLinks = ag.documents.map(
                d => `• ${d.name}: ${PUBLIC_WEB_URL}/sign-director/${encodeURIComponent(d.directors[i].directorSignToken)}`
            );
            const attachments = ag.documents
                .filter(d => d.clientSignedPdfPath && fs.existsSync(d.clientSignedPdfPath))
                .map(d => ({ filename: `ClientSigned_${d.name}.pdf`, path: d.clientSignedPdfPath }));

            await safeSendMail({
                to: email,
                bcc: process.env.AGREEMENTS_INBOX,
                subject: `[AGREEMENT][CLIENT-SIGNED] ${ag.businessName}`,
                text: `Client ${ag.businessName} has completed signatures.\n\nPlease counter-sign:\n\n${directorLinks.join('\n')}`,
                attachments
            });
        }

        res.status(200).json({ complete: true });
    } catch (e) {
        console.error('Client sign error:', e);
        res.status(500).json({ message: e.message || 'Failed to sign document.' });
    }
});

// Debug: directors
app.get('/api/debug/directors-env', (req, res) => {
    res.json({
        DIRECTOR_EMAILS: process.env.DIRECTOR_EMAILS || null,
        DIRECTOR_EMAIL: process.env.DIRECTOR_EMAIL || null,
        parsed: (process.env.DIRECTOR_EMAILS || '').split(',').map(s => s.trim()).filter(Boolean)
    });
});

// Director signs a document
app.post('/api/sign-director/:token', async (req, res) => {
    const { signature, coords } = req.body;
    if (!signature) return res.status(400).json({ message: 'Missing signature' });

    try {
        const decoded = jwt.verify(req.params.token, process.env.JWT_SECRET);
        if (decoded.role !== 'director') return res.status(403).json({ message: 'Wrong link' });

        const { agid, docName, directorIndex = 0 } = decoded;
        const ag = await Agreement.findById(agid);
        if (!ag) return res.status(404).json({ message: 'Agreement not found' });

        const doc = ag.documents.find(d => d.name === docName);
        if (!doc) return res.status(400).json({ message: 'Invalid document token' });
        if (!doc.clientSigned) return res.status(400).json({ message: 'Client must sign first' });
        if (!doc.directors || !doc.directors[directorIndex]) {
            return res.status(400).json({ message: `Director index ${directorIndex} not configured for this document.` });
        }
        const dstate = doc.directors[directorIndex];
        if (dstate.signed) return res.status(400).json({ message: 'Document already signed by you' });

        const template = TEMPLATES.find(t => t.name === doc.name);
        const sigSpec = getDirectorSigSpec(template, directorIndex);
        if (!sigSpec) return res.status(500).json({ message: `No signature coords for director #${directorIndex + 1}` });

        const ts = nowIso().replace(/[:.]/g, '-');
        const safeBusiness = safeFolderName(ag.businessName || 'Client');

        const sigPath = path.join(SIG_DIR, `director${directorIndex + 1}-sig-${doc.name}-${safeBusiness}-${ts}-${ag._id}.png`);
        fs.writeFileSync(sigPath, dataUrlToBuffer(signature));

        const basePdf = (doc.finalSignedPdfPath && fs.existsSync(doc.finalSignedPdfPath))
            ? doc.finalSignedPdfPath
            : doc.clientSignedPdfPath;
        if (!basePdf || !fs.existsSync(basePdf)) return res.status(500).json({ message: `Base PDF missing for ${doc.name}` });

        const outFile = path.join(PDF_DIR, `final-${doc.name}-d${directorIndex + 1}-${ag._id}-${ts}.pdf`);
        await stampSignatureWithTime(basePdf, signature, sigSpec, coords, outFile, `Director ${directorIndex + 1} signed at`);

        dstate.signatureImagePath = sigPath;
        dstate.signature = signature;
        dstate.signed = true;
        dstate.signedDate = new Date();

        doc.finalSignedPdfPath = outFile;
        await ag.save();

        try { await uploadAgreementToSharePoint(ag); }
        catch (e) { console.error('⚠️ SharePoint upload (per director) failed:', e?.response?.data || e.message || e); }

        const nextDoc = ag.documents.find(d =>
            d.clientSigned && d.directors && d.directors[directorIndex] && !d.directors[directorIndex].signed
        );
        if (nextDoc) {
            if (!nextDoc.directors[directorIndex].directorSignToken) {
                nextDoc.directors[directorIndex].directorSignToken = jwt.sign(
                    { agid: ag._id.toString(), role: 'director', docName: nextDoc.name, directorIndex },
                    process.env.JWT_SECRET, { expiresIn: '7d' }
                );
                await ag.save();
            }
            return res.status(200).json({ nextDocToken: nextDoc.directors[directorIndex].directorSignToken });
        }

        const allDirectorsComplete = ag.documents.every(d =>
            d.clientSigned && d.directors && d.directors.length && d.directors.every(x => x.signed)
        );

        if (allDirectorsComplete) {
            const attachments = ag.documents.map(d => ({ filename: `Final_${d.name}.pdf`, path: d.finalSignedPdfPath }));
            await safeSendMail({
                to: ag.email,
                bcc: [...getDirectorEmails(), process.env.AGREEMENTS_INBOX].filter(Boolean).join(','),
                subject: `[AGREEMENT][FINAL] ${ag.businessName}`,
                text: `All documents signed by client & directors. Final PDFs attached.`,
                attachments,
            });

            ag.directorSigned = true;
            ag.directorSignedDate = new Date();
            await ag.save();

            try { await uploadAgreementToSharePoint(ag); }
            catch (e) { console.error('⚠️ Final SharePoint upload failed:', e?.response?.data || e.message || e); }

            return res.status(200).json({ complete: true });
        }

        res.status(200).json({ done: true });
    } catch (e) {
        console.error('Director sign error:', e);
        res.status(500).json({ message: e.message || 'Failed to sign document.' });
    }
});

/* ------------------------- Debug endpoints ------------------------- */
app.get('/api/debug/templates', (req, res) => res.json(listTemplatesStatus()));

app.get('/api/debug/app-env', (req, res) => {
    res.json({
        PUBLIC_WEB_URL: process.env.PUBLIC_WEB_URL || '',
        CORS_ORIGINS: process.env.CORS_ORIGINS || '',
    });
});

app.get('/api/debug/mail-env', (req, res) => {
    res.json({
        MAILER: process.env.SENDGRID_API_KEY ? 'sendgrid' : (process.env.EMAIL_USER && process.env.EMAIL_PASS ? 'gmail' : 'json'),
        EMAIL_USER: !!process.env.EMAIL_USER,
        EMAIL_PASS: process.env.EMAIL_PASS ? '(set)' : null,
        SENDGRID_API_KEY: process.env.SENDGRID_API_KEY ? '(set)' : null,
        MAIL_FROM: process.env.MAIL_FROM || null
    });
});

app.post('/api/debug/mail-test', async (req, res) => {
    const to = req.body.to || process.env.AGREEMENTS_INBOX || process.env.EMAIL_USER;
    try {
        await safeSendMail({
            to,
            subject: 'Mail test from backend',
            text: 'If you see this, your mail transport works.'
        });
        res.json({ ok: true, to });
    } catch (e) {
        res.status(500).json({ ok: false, error: e?.message || String(e) });
    }
});

app.get('/api/health', async (req, res) => {
    const checks = {
        ok: true,
        node_env: process.env.NODE_ENV,
        mongo_uri: !!process.env.MONGO_URI,
        jwt_secret: !!process.env.JWT_SECRET,
        mailer: process.env.SENDGRID_API_KEY ? 'sendgrid' : (process.env.EMAIL_USER && process.env.EMAIL_PASS ? 'gmail' : 'json'),
        templates: listTemplatesStatus()
    };
    try { await mongoose.connection.db.admin().ping(); checks.mongo = 'ok'; }
    catch { checks.mongo = 'not_connected'; }
    res.json(checks);
});

/* ------------------------- Root & start ------------------------- */
app.get('/', (req, res) => res.send('Backend is live 🚀'));
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));
