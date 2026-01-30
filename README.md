# 📊 NIFTY Volatility & Sigma Analysis Dashboard

## Overview
This project is a **volatility-driven analytical dashboard** built exclusively for the **NIFTY index**, designed to quantify **expected price ranges**, **probability of movement**, and **market normality** using **India VIX**, realised volatility, and historical distribution analysis.

The system converts volatility into **actionable range expectations** across **Daily, Weekly, and Monthly** horizons, enabling probability-based trade structuring rather than directional guessing.

![NIFTY Volatility & Sigma Dashboard](Images/Dashboard.png)

---

## Core Objectives
- Quantify **expected NIFTY price movement** using volatility
- Measure **confidence levels (probabilities)** for price ranges
- Detect **normal vs abnormal market behavior** in real time
- Support **range selection & strike placement** with statistical backing

---

## Key Components

### 1. Sigma-Based Expected Move Engine
- Computes **±1σ expected move (68% confidence)** for:
  - Daily
  - Weekly (current expiry)
  - Monthly (current expiry)
- Sigma is derived from **current India VIX**, adjusted for time to expiry (DTE).
- Output is expressed as **percentage return ranges**, not absolute bias.

---

### 2. Live Normal Distribution Monitor
A continuously updating **normal distribution curve** tracks real-time NIFTY movement.

**Purpose**
- Visually assess whether price action is:
  - Within expected volatility bounds
  - Approaching tail risk
  - Breaking statistical assumptions

**Live overlays**
- ±1σ, ±2σ, ±3σ bands
- Current price position on the curve
- Day’s realised high / low relative to distribution

This allows instant judgment on whether the market is behaving **statistically normal or stressed**.

---

### 3. Volatility Intelligence Panel
Displayed alongside the distribution:

- India VIX (current)
- VIX % change from previous close
- Intraday VIX high / low deviation
- Realised Volatility:
  - 20 trading sessions
  - 30 trading sessions
- IV Percentile (IVP)

This contextualizes **implied vs realised volatility** and highlights volatility mispricing.

---

### 4. NIFTY vs India VIX Time-Series Tracker
A synchronized chart showing:
- NIFTY index movement
- India VIX movement

**Use case**
- Identify divergence / convergence
- Detect volatility compression or expansion phases
- Validate whether price movement is volatility-supported

---

### 5. Historical Range Distribution Analysis (3 Years)
A statistical study of **actual NIFTY movement** over the past **3 years**, covering:
- Daily
- Weekly
- Monthly

For each timeframe:
- Tracks **realised high–low range**
- Builds **percentile distributions**
- Answers questions like:
  - “What range contained price 75% of the time?”
  - “How often did NIFTY exceed 1% in a day?”

This creates **empirical confidence bands**, independent of VIX assumptions.

---

### 6. Probability-Based Range Selection
By combining:
- Historical percentiles
- Current volatility regime
- Time horizon

The system identifies **statistically favorable ranges** for:
- Option selling
- Iron condors
- Strangles / straddles
- Range-bound strategies

The focus is on **Probability of Profit (PoP)**, not prediction.

---

### 7. Forward Range Projection (Custom Days)
Using the **current India VIX level**, the system allows the user to input **any number of future trading days** (not limited to standard expiries).

For the selected horizon, the model computes:
- Expected price movement ranges
- **Confidence levels expressed as percentiles (1%–100%)**

The percentile framework represents the **historical probability** that NIFTY’s price remains within a given range for the specified period, based on volatility scaling and historical distribution behavior.

**Key characteristics**
- Confidence is not fixed to standard levels (e.g. 68% or 95%)
- Any percentile can be evaluated (e.g. 60%, 72%, 85%, 93%)
- Ranges expand monotonically with higher confidence

**Practical applications**
- Fine-tuned range selection for positional trades
- Volatility-aware hedge sizing
- Event-specific risk assessment where custom time horizons are required

This approach enables **continuous probability modeling** rather than discrete confidence buckets, supporting more precise risk and payoff planning.

---

## Design Philosophy
- Instrument-specific (NIFTY only)
- Volatility-first, direction-agnostic
- Probability over opinion
- Visual + statistical confirmation
- Practical for live trading decisions

---

## Intended Users
- Options traders
- Volatility traders
- Quantitative analysts
- Systematic strategy developers

---

## Disclaimer
This tool is for **analytical and educational purposes**.  
It does not constitute financial advice or trade recommendations.

---

