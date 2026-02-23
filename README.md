# 📄 Local-AI-MS-Word-Report-Generator

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)
[![LM Studio](https://img.shields.io/badge/Integration-LM%20Studio-orange.svg)](https://lmstudio.ai/)

<div align="center">
  <h3>Transform simple titles into professional technical reports using local AI</h3>
  <p><em>No API keys • No internet required • 100% private • Runs on your machine</em></p>
</div>

---

## 📋 Overview

**AI Report Generator** is a PowerShell script with a Windows Forms GUI that connects to a local LM Studio instance to automatically generate complete, professionally formatted Microsoft Word reports. Simply enter a title and detailed instructions, and the AI handles everything:

- 📝 **Generates** a structured outline based on your specific requirements
- ✍️ **Creates** 3-5 detailed paragraphs for each section
- 📊 **Builds** a properly formatted Word document with headings and professional layout
- 💾 **Saves** the finished report to your desktop

Perfect for technical reports, research papers, proposals, case studies, and analytical documents.

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🖥️ **Native Windows GUI** | Clean, intuitive interface with large instructions textbox |
| 🤖 **Local AI Processing** | Works with LM Studio – completely offline and private |
| 📝 **Paragraph Generation** | Creates 3-5 detailed paragraphs per section (not just bullets) |
| 📋 **Custom Instructions** | Large multiline textbox for detailed guidance to the AI |
| 📊 **Microsoft Word Integration** | Full COM automation with proper heading styles and formatting |
| 🔧 **Configurable Timeouts** | Adjustable settings for slower systems |
| 🔄 **Automatic Fallbacks** | Works even if AI service is unavailable |
| 📁 **Auto-save** | Timestamped files saved to desktop |

---

## 🚀 Quick Start

### Prerequisites

1. **LM Studio** installed and running with server mode enabled
2. **Microsoft Word** (2013 or later recommended)
3. **PowerShell 5.1+** (comes pre-installed with Windows 10/11)
4. **Model loaded** in LM Studio (tested with `liquid/lfm2-1.2b` and similar models)

### Installation

```powershell
# Clone the repository
# Navigate to directory
# Modify .ps1 with your own ip, model
# Run the script (you may need to bypass execution policy)
