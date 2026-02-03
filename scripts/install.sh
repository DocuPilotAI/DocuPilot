#!/usr/bin/env bash
# DocuPilot One-Click Installation Script
# Usage: curl -fsSL https://raw.githubusercontent.com/docupilot/docupilot/main/scripts/install.sh | bash
# Or:    bash -c "$(curl -fsSL https://raw.githubusercontent.com/docupilot/docupilot/main/scripts/install.sh)"
#
# Custom options:
#   DOCUPILOT_REPO - Git repository URL (default: https://github.com/docupilot/docupilot.git)
#   DOCUPILOT_BRANCH - Branch name (default: main)
#   DOCUPILOT_DIR - Installation directory (default: ./docupilot)

set -e

REPO="${DOCUPILOT_REPO:-https://github.com/docupilot/docupilot.git}"
BRANCH="${DOCUPILOT_BRANCH:-main}"
DIR="${DOCUPILOT_DIR:-./docupilot}"

echo ""
echo "╔════════════════════════════════════╗"
echo "║  DocuPilot Installation Script     ║"
echo "╚════════════════════════════════════╝"
echo ""

# Check required commands
echo "🔍 Checking system environment..."

if ! command -v git &> /dev/null; then
  echo "❌ git not found, please install it first:"
  echo ""
  if [[ "$OSTYPE" == "darwin"* ]]; then
    echo "  macOS: xcode-select --install"
  elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
    echo "  Linux: sudo apt-get install git  (or use your package manager)"
  fi
  exit 1
fi

if ! command -v node &> /dev/null; then
  echo "❌ Node.js not found, please install Node.js 18.0 or higher:"
  echo ""
  echo "  Visit: https://nodejs.org/"
  exit 1
fi

if ! command -v npm &> /dev/null; then
  echo "❌ npm not found, please ensure Node.js is properly installed"
  exit 1
fi

NODE_VERSION=$(node -v | cut -d'v' -f2 | cut -d'.' -f1)
if [ "$NODE_VERSION" -lt 18 ]; then
  echo "⚠️  Node.js version too low (current: $(node -v), required: >= 18.0.0)"
  echo "Please upgrade Node.js: https://nodejs.org/"
  exit 1
fi

echo "✅ git: $(git --version)"
echo "✅ node: $(node -v)"
echo "✅ npm: $(npm -v)"
echo ""

echo "📦 Installation configuration:"
echo "   Repository: $REPO"
echo "   Branch: $BRANCH"
echo "   Directory: $DIR"
echo ""

# Check directory
if [ -d "$DIR" ]; then
  if [ -n "$(ls -A $DIR 2>/dev/null)" ]; then
    echo "⚠️  Directory $DIR already exists and is not empty"
    read -p "Delete and reinstall? (y/N): " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
      echo "🗑️  Deleting existing directory..."
      rm -rf "$DIR"
    else
      echo "❌ Installation cancelled"
      echo "Tip: Set environment variable DOCUPILOT_DIR=./my-dir to specify another directory"
      exit 1
    fi
  fi
fi

# Clone repository
echo "📥 Cloning repository (branch: $BRANCH)..."
if ! git clone --depth 1 --branch "$BRANCH" "$REPO" "$DIR"; then
  echo "❌ Clone failed, please check network connection or repository URL"
  exit 1
fi

cd "$DIR"

# Detect project structure
if [ -d "DocuPilot" ]; then
  echo "📂 Detected subdirectory structure, entering DocuPilot/"
  cd DocuPilot
  APP_DIR="$DIR/DocuPilot"
else
  APP_DIR="$DIR"
fi

# Install dependencies
echo "📦 Installing npm dependencies..."
if ! npm install; then
  echo "❌ Dependencies installation failed"
  exit 1
fi

# Configure environment variables
if [ ! -f ".env.local" ]; then
  if [ -f ".env.local.example" ]; then
    echo "⚙️  Creating .env.local configuration file..."
    cp .env.local.example .env.local
    echo "✅ Created .env.local, please edit and add your Anthropic API Key"
  fi
fi

echo ""
echo "╔════════════════════════════════════╗"
echo "║   ✅ Installation Successful!      ║"
echo "╚════════════════════════════════════╝"
echo ""
echo "📍 Installation location: $APP_DIR"
echo ""
echo "🚀 Next steps:"
echo ""
echo "1️⃣  Enter project directory:"
echo "   cd $APP_DIR"
echo ""
echo "2️⃣  Configure API Key (if not already configured):"
echo "   Edit .env.local file and add your Anthropic API Key"
echo "   ANTHROPIC_API_KEY=sk-ant-your-key-here"
echo ""
echo "3️⃣  Start development server:"
echo "   npm run dev:https"
echo ""
echo "4️⃣  Load add-in in Office:"
echo "   - Open Excel/Word/PowerPoint"
echo "   - Insert > My Add-ins > Upload My Add-in"
echo "   - Select: $APP_DIR/manifest.xml"
echo ""
echo "📖 Full documentation: https://github.com/docupilot/docupilot"
echo ""
