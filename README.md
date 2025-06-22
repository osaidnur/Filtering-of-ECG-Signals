# 🫀 ECG Signal Filtering Project

## 📕 Table of Contents

- [📋 Project Overview](#-project-overview)
- [🎯 Project Objectives](#-project-objectives)
- [🛠️ Technical Implementation](#️-technical-implementation)
- [🚀 Getting Started](#-getting-started)
- [📋 Menu Options](#-menu-options)
- [📁 Project Structure](#-project-structure)
- [🔬 Signal Processing Theory](#-signal-processing-theory)
- [📊 Results & Analysis](#-results--analysis)


## 📋 Project Overview

This project implements digital signal processing techniques to filter real electrocardiographic (ECG) signals for improved readability and automatic processing. The system applies high-pass and low-pass filters to remove noise and artifacts from ECG recordings, making them suitable for subsequent peak detection and classification tasks.

### 👥 Team Members
- **Osaid Nur** - 1210733
- **Salah Sami** - 1210722  
- **Waleed Rimawi** - 1211491

## 🎯 Project Objectives

The main goals of this project are to:

1. **Data Visualization**: Display and analyze real ECG signals sampled at 360 Hz
2. **Signal Filtering**: Apply digital filters to improve signal quality
3. **Filter Analysis**: Analyze filter characteristics including frequency response and zero-pole plots
4. **Comparative Analysis**: Compare different filtering approaches and their effects

## 📊 Features

### 🔍 Signal Analysis
- **Original Signal Display**: Visualize raw ECG1 and ECG2 signals
- **Real-time Column Integration**: Time axis generation for proper signal representation
- **Interactive Plotting**: High-quality matplotlib visualizations with customizable parameters

### 🔧 Digital Filtering
- **High-Pass Filter (HPF)**: Removes baseline drift and low-frequency noise
- **Low-Pass Filter (LPF)**: Eliminates high-frequency artifacts and muscle noise
- **Cascaded Filtering**: Apply multiple filters in sequence for enhanced signal quality

### 📈 Filter Characterization
- **Frequency Response Analysis**: Visualize filter magnitude response
- **Zero-Pole Plot Generation**: Analyze filter stability and characteristics
- **Transfer Function Computation**: Mathematical representation of filter behavior

## 🛠️ Technical Implementation

### High-Pass Filter Specifications
```
Transfer Function: H_HP(z) = (-1/32 + z^-16 - z^-17 + z^-32/32) / (1 - z^-1)
Type: IIR (Infinite Impulse Response)
Purpose: Baseline drift removal, DC component elimination
```

### Low-Pass Filter Specifications  
```
Transfer Function: H_LP(z) = (1 - 2z^-6 + z^-12) / (1 - 2z^-1 + z^-2)
Type: IIR (Infinite Impulse Response)  
Purpose: High-frequency noise suppression, anti-aliasing
```

## 🚀 Getting Started

### Prerequisites
```bash
pip install numpy matplotlib pandas scipy
```

### Required Files
- `Data_ECG_raw.xlsx` - Contains ECG1 and ECG2 signal data
- `DSP_project.py` - Main application script

### Installation & Usage

1. **Clone the repository:**
```bash
git clone https://github.com/yourusername/Filtering-of-ECG-Signals.git
cd Filtering-of-ECG-Signals
```

2. **Install dependencies:**
```bash
pip install -r requirements.txt
```

3. **Run the application:**
```bash
python DSP_project.py
```

## 📋 Menu Options

The interactive application provides the following options:

1. **Display Original Signals** 📊
   - View raw ECG1 and ECG2 recordings
   - Analyze signal characteristics before filtering

2. **Filter Frequency Response** 📈
   - Visualize HPF and LPF magnitude responses
   - Understand filter behavior in frequency domain

3. **Zero-Pole Analysis** 🎯
   - Plot filter zeros and poles
   - Assess filter stability and performance

4. **Single Filter Application** 🔧
   - Apply HPF or LPF to individual signals
   - Compare filtered vs. original signals

5. **Cascaded Filter Analysis** 🔄
   - Apply HPF→LPF or LPF→HPF sequences
   - Analyze order-dependent filtering effects

## 📁 Project Structure

```
Filtering-of-ECG-Signals/
├── DSP_project.py          # Main application script
├── Data_ECG_raw.xlsx       # Raw ECG signal data
├── Project Report.pdf      # Detailed technical report
├── Project_Description.pdf # Project requirements and specifications
├── README.md              # Project documentation
└── requirements.txt       # Python dependencies
```

## 🔬 Signal Processing Theory

### ECG Signal Characteristics
- **Sampling Rate**: 360 Hz
- **Signal Components**: 
  - P-wave, QRS complex, T-wave
  - Baseline drift (< 0.5 Hz)
  - Muscle artifacts (> 100 Hz)
  - Power line interference (50/60 Hz)

### Filtering Strategy
1. **High-Pass Filtering**: Removes baseline wander and respiratory artifacts
2. **Low-Pass Filtering**: Eliminates high-frequency noise and EMG interference
3. **Cascaded Approach**: Combines both filters for optimal signal conditioning

## 📊 Results & Analysis

### Filter Performance
- **HPF Effectiveness**: Successfully removes baseline drift while preserving ECG morphology
- **LPF Characteristics**: Maintains clinical ECG features while suppressing noise
- **Order Independence**: HPF→LPF vs. LPF→HPF comparison reveals filtering sequence effects

### Signal Quality Metrics
- **SNR Improvement**: Quantifiable enhancement in signal-to-noise ratio
- **Morphology Preservation**: Clinical ECG features remain intact
- **Computational Efficiency**: Real-time processing capability

---