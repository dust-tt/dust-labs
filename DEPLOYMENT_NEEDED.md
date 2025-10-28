# PowerPoint Add-in Production Deployment Needed

## User Issue

**Reported:** October 9, 2025
**Problem:** PowerPoint add-in shows `InvalidArgument` error when processing slides

## Root Cause

Production code tries to read text from all shapes, including images and charts.
Images and charts don't have text properties, so the code fails.

## The Fix

The bug was fixed on September 11, 2025 (commits f7705d7, c7c1846, d50570c).
The fix is in the repository but hasn't been deployed to production yet.

## What the Fix Does

**Before (production - broken):**
- Tries to load text properties from all shapes
- Crashes when it hits an image or chart

**After (repository - fixed):**
- Checks shape type first
- Only loads text from text-capable shapes (text boxes, titles, etc.)
- Handles errors gracefully

## Next Step

Deploy the current main branch to production:

```bash
cd /Users/iliasbettahi/dust-labs/office
vercel --prod
```

## Testing

After deployment, test with a PowerPoint containing:
- Text boxes ✅
- Images 🖼️
- Charts 📊

All should work without errors.

---

**Issue source:** User ticket from Guilhem Delorme (workspace M1K8WmfYjN)
**Fix commits:** f7705d7, c7c1846, d50570c, a8c0a62, 3f563c2 (Sept 11, 2025)
**Action:** Deploy to production
