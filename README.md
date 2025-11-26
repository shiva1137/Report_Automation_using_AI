# 🤖 FSA Report Automation using AI

> An intelligent Telegram bot that automates FSA (Filling Station Automation) trip data report generation using Natural Language Processing (NLP) and AI-powered query understanding.

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Telegram](https://img.shields.io/badge/Telegram-Bot-blue.svg)](https://telegram.org/)

## 📋 Table of Contents

- [Overview](#-overview)
- [Features](#-features)
- [Tech Stack](#-tech-stack)
- [Architecture](#-architecture)
- [Installation](#-installation)
- [Configuration](#-configuration)
- [Usage](#-usage)
- [Project Structure](#-project-structure)
- [Key Features Explained](#-key-features-explained)
- [Performance Optimizations](#-performance-optimizations)
- [Contributing](#-contributing)

## 🎯 Overview

This project is an **AI-powered Telegram bot** that enables users to generate Excel reports for FSA trip data through natural language queries. Instead of writing complex database queries or using traditional forms, users can simply ask the bot in plain English:

> *"Give me Excel file for PS trips for Area 1 for Jun 2024"*

The bot uses **GPT-4o** (via LLM7.io) to understand natural language, extract structured information, query MongoDB, and automatically generate formatted Excel reports.

### Problem Solved

- **Before**: Users needed technical knowledge to query databases, write SQL/MongoDB queries, and format reports manually
- **After**: Users can request reports in natural language via Telegram, and the bot handles everything automatically

## ✨ Features

### 🤖 Natural Language Processing
- **Intelligent Query Parsing**: Understands natural language queries using GPT-4o
- **Multiple Query Formats**: Supports various date formats, categories, and area specifications
- **Conversation Flow**: Bot asks for missing information interactively
- **Context Understanding**: Handles ambiguous queries and requests clarification

### 📊 Flexible Data Queries
- **Multiple Categories**: Request single or multiple trip categories (PS, MC, JR, DFW)
- **Multiple Areas**: Query single area, multiple areas, or all 15 areas
- **Date Ranges**: Support for single months, date ranges, month-only, and year-only queries
- **All Combinations**: Request all categories and all areas simultaneously

### 📈 Excel Report Generation
- **Automated Formatting**: Professional Excel files with borders, alignment, and formatting
- **Dynamic Filenames**: Auto-generated filenames based on area, category, and date
- **Memory Efficient**: Handles large datasets with optimized memory management

### 🔧 Performance & Reliability
- **Connection Pooling**: MongoDB connection reuse with automatic idle timeout
- **Memory Management**: Automatic garbage collection and resource cleanup
- **Retry Logic**: Automatic retries for network and database operations
- **Error Handling**: Comprehensive error handling with detailed logging

### 🚀 Auto-Dependency Management
- **Automatic Installation**: Checks and installs missing packages from `requirements.txt`
- **Zero-Configuration Setup**: Just run the script and it handles dependencies

## 🛠 Tech Stack

### Core Technologies
- **Python 3.8+**: Main programming language
- **MongoDB Atlas**: Cloud database for trip data storage
- **Telegram Bot API**: User interface and communication
- **LLM7.io (GPT-4o)**: Natural Language Processing and AI understanding

### Key Libraries
- `python-telegram-bot`: Telegram bot framework
- `motor`: Async MongoDB driver
- `pandas`: Data manipulation and Excel generation
- `openpyxl`: Excel file formatting
- `openai`: LLM API client (via LLM7.io)
- `tenacity`: Retry logic for reliability
- `pytz`: Timezone handling

## 🏗 Architecture

```
┌─────────────────────────────────────────────────────────┐
│              Telegram Bot (User Interface)              │
│  - Receives natural language queries                    │
│  - Sends Excel reports                                  │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│            NLP Query Parser (GPT-4o)                     │
│  - Extracts categories, areas, dates                     │
│  - Handles conversation flow                             │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│         MongoDB Connection Manager                      │
│  - Connection pooling                                   │
│  - Auto-close idle connections                          │
└────────────────────┬────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────┐
│            Data Processing Engine                        │
│  - Fetches trip data                                    │
│  - Filters by area/category/date                         │
│  - Generates Excel reports                              │
└─────────────────────────────────────────────────────────┘
```

### Data Flow

1. **User Query** → Telegram Bot receives natural language query
2. **NLP Parsing** → GPT-4o extracts structured information (categories, areas, dates)
3. **Conversation** → Bot asks for missing information if needed
4. **Data Fetching** → Queries MongoDB with extracted parameters
5. **Report Generation** → Creates formatted Excel file
6. **Delivery** → Sends Excel file to user via Telegram

## 📦 Installation

### Prerequisites
- Python 3.8 or higher
- MongoDB Atlas account (or MongoDB instance)
- Telegram Bot Token (from [@BotFather](https://t.me/botfather))
- LLM7.io API key

### Step 1: Clone the Repository
```bash
git clone https://github.com/yourusername/FSA_Report_Automation_using_AI.git
cd FSA_Report_Automation_using_AI
```

### Step 2: Install Dependencies
The script automatically checks and installs dependencies, but you can also install manually:

```bash
pip install -r requirements.txt
```

### Step 3: Configure Environment Variables

**Option 1: Using .env file (Recommended)**

1. Create a `.env` file in the project root:
```bash
cp .env.example .env
```

2. Edit `.env` and fill in your credentials:
```env
MONGO_CONNECTION_STRING=your_mongodb_connection_string_here
TELEGRAM_BOT_TOKEN=your_telegram_bot_token_here
TELEGRAM_CHAT_ID=your_chat_id_here
LLM7_API_KEY=your_llm7_api_key_here
```

**Option 2: Using Environment Variables**

**Windows (PowerShell):**
```powershell
$env:MONGO_CONNECTION_STRING="mongodb://user:pass@host:port/db"
$env:TELEGRAM_BOT_TOKEN="your-bot-token"
$env:TELEGRAM_CHAT_ID="-123456789"
$env:LLM7_API_KEY="your-api-key"
```

**Windows (CMD):**
```cmd
set MONGO_CONNECTION_STRING=mongodb://user:pass@host:port/db
set TELEGRAM_BOT_TOKEN=your-bot-token
set TELEGRAM_CHAT_ID=-123456789
set LLM7_API_KEY=your-api-key
```

**Linux/Mac:**
```bash
export MONGO_CONNECTION_STRING="mongodb://user:pass@host:port/db"
export TELEGRAM_BOT_TOKEN="your-bot-token"
export TELEGRAM_CHAT_ID="-123456789"
export LLM7_API_KEY="your-api-key"
```

### Step 4: Get Your Credentials

1. **MongoDB Connection String**: Get from your MongoDB Atlas dashboard or MongoDB instance
2. **Telegram Bot Token**: Create a bot with [@BotFather](https://t.me/botfather) on Telegram
3. **Telegram Chat ID**: Message [@userinfobot](https://t.me/userinfobot) to get your chat ID
4. **LLM7.io API Key**: Sign up at [llm7.io](https://llm7.io) and get your API key

> 🔒 **Security Note**: 
> - **NEVER commit `.env` file or hardcode credentials in the script**
> - All sensitive data is loaded from environment variables
> - The `.env` file is already in `.gitignore` to prevent accidental commits
> - This ensures your production credentials stay secure

### Step 5: Run the Bot
```bash
python FSA_Report_Automation_using_AI.py
```

## ⚙️ Configuration

### Available Categories
- `MC`: Motorcycle trips
- `JR`: Junior trips
- `PS`: Petrol Station trips
- `DFW`: Diesel Filling trips

### Available Areas
15 areas covering different regions (Area-1 through Area-15)

### Performance Settings
- `MAX_WORKERS`: Maximum concurrent database queries (default: 500)
- `MONGO_IDLE_TIMEOUT`: Connection idle timeout in seconds (default: 300)

## 💬 Usage

### Basic Queries

**Single Category, Single Area:**
```
Give me Excel file for PS trips for Area 1 for Jun 2024
```

**Multiple Categories:**
```
PS and MC trips Area 1 Jun 2024
```

**Multiple Areas:**
```
PS trips for Area 1 and Area 2 Jun 2024
```

**Date Ranges:**
```
MC trips Area 5 Jun 2024 to Aug 2024
```

**All Categories:**
```
All categories Area 1 Jun 2024
```

**All Areas:**
```
All areas for Jun 2024
```

**Month Only (finds last occurrence):**
```
August trips
```

### Supported Date Formats
- ✅ `"Jun 2024"`, `"June 2024"`
- ✅ `"Jun 2024 to Aug 2024"` (date ranges)
- ✅ `"August"` or `"Aug"` (month only - finds last occurrence)
- ✅ `"2024"` (full year)
- ✅ `"Jan 2025"`, `"January 2025"`

### Conversation Flow
If information is missing, the bot will ask:

```
User: "Give me PS trips"
Bot: "For what period would you like the Excel file?"
User: "Jun 2024"
Bot: "For which area(s) would you like the Excel file?"
User: "Area 1"
Bot: [Processes and sends Excel file]
```

## 📁 Project Structure

```
FSA_Report_Automation_using_AI/
│
├── FSA_Report_Automation_using_AI.py  # Main bot script
├── requirements.txt                   # Python dependencies
├── README.md                          # This file
├── .env.example                       # Example environment variables (template)
├── .env                               # Your actual credentials (NOT in git)
└── .gitignore                         # Git ignore rules
```

> ⚠️ **Important**: The `.env` file contains your actual credentials and is **NOT** tracked by git. Always use `.env.example` as a template.

## 🔑 Key Features Explained

### 1. Automatic Dependency Management
The script automatically checks for required packages and installs missing ones from `requirements.txt` before running.

### 2. MongoDB Connection Manager
- **Singleton Pattern**: Single global instance manages all connections
- **Connection Reuse**: Reuses existing connections across queries
- **Auto-Close**: Automatically closes idle connections after 5 minutes
- **Memory Efficient**: Triggers garbage collection after connection closures

### 3. NLP Query Parsing
Uses GPT-4o to extract:
- Categories (single, multiple, or all)
- Areas (single, multiple, or all)
- Date/period information (various formats)

### 4. Memory Optimization
- Automatic garbage collection after memory-intensive operations
- Temporary Excel files cleaned up after sending
- DataFrames deleted immediately after use
- Connection pooling reduces memory overhead

### 5. Error Handling & Retry Logic
- Automatic retries for MongoDB queries (3 attempts with exponential backoff)
- Automatic retries for Telegram API calls
- Comprehensive error logging to file and console

## ⚡ Performance Optimizations

### Memory Management
- ✅ Automatic garbage collection after large operations
- ✅ DataFrame cleanup immediately after use
- ✅ Temporary file removal after sending
- ✅ Connection pooling with idle timeout

### Database Optimization
- ✅ Parallel batch processing (up to 500 concurrent queries)
- ✅ Efficient MongoDB aggregation pipelines
- ✅ Connection reuse across queries
- ✅ Automatic connection cleanup

### Response Time
- **Typical Response**: 5-30 seconds per query
- **Factors**: Data volume, number of categories/areas, date range size

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🔒 Security & Privacy

### For Portfolio/Public Repository

This project is designed to be **safe for public GitHub repositories**:

✅ **All sensitive credentials are loaded from environment variables**  
✅ **No hardcoded passwords, tokens, or connection strings**  
✅ **`.env` file is in `.gitignore` to prevent accidental commits**  
✅ **Configuration validation ensures required variables are set**

### Before Pushing to GitHub

1. ✅ Verify no credentials are hardcoded in the code
2. ✅ Ensure `.env` is in `.gitignore` (already included)
3. ✅ Use `.env.example` as a template for others
4. ✅ Never commit production credentials

### For Your Office/Production Use

- Keep your actual `.env` file **local only**
- Use different credentials for demo/portfolio vs production
- Consider using MongoDB Atlas free tier for demo purposes
- Create a separate Telegram bot for testing

## 👤 Author

**Your Name**
- GitHub: [@yourusername](https://github.com/yourusername)
- LinkedIn: [Your LinkedIn](https://linkedin.com/in/yourprofile)

## 🙏 Acknowledgments

- [python-telegram-bot](https://github.com/python-telegram-bot/python-telegram-bot) - Telegram bot framework
- [LLM7.io](https://llm7.io) - AI/LLM API provider
- [MongoDB Atlas](https://www.mongodb.com/cloud/atlas) - Cloud database service

---

⭐ If you find this project helpful, please consider giving it a star!

