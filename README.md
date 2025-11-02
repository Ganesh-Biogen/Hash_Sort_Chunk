# 🧮 Deterministic Hash–Sort–Chunk Sampling (HSC)

[![DOI](https://img.shields.io/badge/DOI-10.13140/RG.2.2.27435.50721-blue.svg)](https://doi.org/10.13140/RG.2.2.27435.50721)

**Author:** Ganesh Raj  
**Date:** 2025  
**Domain:** Data Comparison, Database Consistency Validation, Data Engineering  

---

## 📘 Overview

**Hash–Sort–Chunk Sampling (HSC)** is a deterministic, statistically principled method designed to efficiently compare large datasets across **heterogeneous databases**.

The technique uses:
1. **Hashing** — Converts record keys into uniformly distributed hash values (e.g., SHA-256).  
2. **Sorting** — Sorts records by their hashes to destroy clustering and ensure uniformity.  
3. **Chunking & Sampling** — Divides the dataset into equal-sized chunks and samples deterministically for efficient difference detection.

Unlike random sampling or Merkle Trees, HSC:
- Offers **consistent reproducibility** (same results across runs).  
- Ensures **uniform coverage** across the entire dataset.  
- Detects differences even in **distributed or dense mismatch scenarios**.  
- Is **database-agnostic**, suitable for one-time validation between heterogeneous systems.

---

## ⚙️ Methodology Summary

### Step 1 — Hashing
Each record’s primary/composite key is hashed:
