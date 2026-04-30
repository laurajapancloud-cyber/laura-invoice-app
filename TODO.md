# 🎯 Mission: LAURA JAPAN PWA + Async Job Queue Implementation

## 📂 Phase 1: PWA Implementation
- [ ] Create `manifest.json` with luxury theme
- [ ] Create `sw.js` (Cache-first for static, Network-first for API)
- [ ] Generate PWA Icons (192x192, 512x512, Maskable, Apple)
- [ ] Update `main.py` to serve PWA files with correct MIME types
- [ ] Update `index.html` <head> with PWA meta tags
- [ ] Implement SW registration and Custom Install UI in `index.html`
- [ ] Verification: Lighthouse PWA score >= 90

## 🔥 Phase 2: Async Job Queue
### 2-A. Server-side Job API
- [ ] Implement `jobs` table migration in `main.py`
- [ ] Refactor `analyze_images` and `upload_to_drive` into internal shared functions
- [ ] Add `POST /api/jobs/analyze` and `POST /api/jobs/drive-upload`
- [ ] Add `GET /api/jobs/{id}` and list endpoints

### 2-B. Frontend: Persistence & Manager
- [ ] Implement IndexedDB wrapper in `index.html`
- [ ] Implement `JobManager` for polling and lifecycle
- [ ] Implement `JobUI` for the bottom-right panel
- [ ] Handle persistence (Restore polling on page reload)

### 2-C. UI & Integration
- [ ] Add `#job-queue` panel to `index.html` with luxury styling
- [ ] Replace existing Analyze/Drive buttons with Job Queue enqueuing
- [ ] Handle completion callbacks (Inyect results into UI)

### 2-D. Final Verification
- [ ] Scenario A: Tab switching during job
- [ ] Scenario B: Parallel jobs
- [ ] Scenario C: Persistence (Reload)
- [ ] Scenario D: API backward compatibility

## 📊 Final Artifacts
- [ ] `IMPLEMENTATION_REPORT.md`
- [ ] E2E Test Screenshots/Videos
