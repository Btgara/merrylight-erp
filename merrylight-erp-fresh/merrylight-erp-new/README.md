# 🍓 Merrylight Packhouse ERP

Complete Enterprise Resource Planning system for Merrylight Enterprises packhouse operations.

## ✨ FEATURES

### 📊 Dashboard
- Real-time statistics
- Daily metrics

### 🍓 Fruit Receiving
- Batch upload & tracking
- Net weight calculation
- Barcode printing (CODE128)
- Advanced filtering

### 📦 Packing Operations (5 Tabs)
1. **Work Order Creation** - Create customer orders (Customer, Variety, Pack Type from Master Files)
2. **Tip to Work Order** - Link fruit batches to orders
3. **Pallet Registration** - Register and track pallets
4. **Pallets Report** - View all pallets
5. **Packout Summary** - Operations overview with metrics

### ✅ Quality Control
- QC checks
- Defect tracking

### 👥 HR Management
- Employee records
- Contract management

### ⚙️ Master Files (15 Types)
- Customers, Varieties, Pack Types (used in Work Orders)
- Sections, Blocks, Commodities
- Pallet Types, Crate Types, Grades
- Seasons, Channels, Transport, Sizes, Brands, Invoice Codes

## 🚀 QUICK START

### 1. Install Dependencies
```bash
npm install
```

### 2. Start Locally
```bash
npm start
```

### 3. Build for Production
```bash
npm run build
```

### 4. Deploy on Vercel
- Push to GitHub
- Connect to Vercel
- Auto-deploys! 🎉

## 🔐 ROLES

- **Admin** - Full access
- **HR Manager** - Most operations
- **Entry/QC** - Receiving, Packing, QC
- **Viewer** - Read-only

## 💾 DATA

Browser localStorage (no database needed)
- Data persists locally
- Per browser/device

## 📱 BROWSER SUPPORT

- Chrome, Firefox, Safari, Edge (latest)

## 🎯 WORKFLOW

1. Set up Master Files (Customers, Varieties, Pack Types)
2. Create fruit batches (Fruit Receiving)
3. Create work orders (Packing Operations)
4. Link fruit to orders
5. Build pallets
6. Track operations

## 📝 VERSION

1.0.0 - Production Ready ✅

---

**Created:** July 2026
**Status:** ✅ Production Ready
**Deployment:** Vercel

See QUICK_START.md for deployment guide!
