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
Each record’s primary/composite key is hashed:hash = SHA256(key_fields)

### Step 2 — Sorting
All records are sorted by their hash values to remove spatial correlation:dataset.sort(by='hash')

### Step 3 — Chunking & Sampling
The sorted dataset is split into *C* equal chunks.  
A deterministic sample is taken per chunk:
sample_size = 1
sample = [chunk[i % len(chunk)] for i, chunk in enumerate(chunks)]


The probability that a differing row falls into a given chunk follows a **Binomial distribution**:
\[
P(k) = \binom{D}{k} \left(\frac{1}{C}\right)^k \left(1 - \frac{1}{C}\right)^{D-k}
\]

---

## 📊 Advantages

| Method | Strength | Limitation |
|--------|-----------|-------------|
| **Checksum / Hash Aggregation** | Detects if differences exist | Cannot locate where or what differs |
| **Merkle Trees** | Efficient for clustered/sparse differences | Expensive for distributed differences |
| **Bloom Filters** | Fast membership checks | No field-level diagnostics |
| **HSC (Proposed)** | Uniform coverage, deterministic, diagnostic | Requires initial sort and hash pass |

---
##    📊 Citation
If you use this method or code, please cite:

Ganesh Raj. Deterministic Hash–Sort–Chunk Sampling for Efficient Database Comparison. ResearchGate, 2025.
DOI: 10.13140/RG.2.2.27435.50721

---
##    📬 Contact

For discussions, improvements, or collaboration:
Author: Ganesh Raj
LinkedIn / ResearchGate: [ResearchGate Profile](https://www.researchgate.net/profile/Ganesh-Raj-Munikrishnan?enrichId=rgreq-4c2f730cd8f31e9afb3ee4abaa0cf7a8-XXX&enrichSource=Y292ZXJQYWdlOzM5NzE3OTA2NTtBUzoxMTQzMTI4MTcwOTc0MjUyMEAxNzYyMDc5NjgzODYx&el=1_x_10&_esc=publicationCoverPdf)
---
