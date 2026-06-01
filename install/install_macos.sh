#!/usr/bin/env bash
# ============================================================
# MinervaDB XL View - macOS Automated Installer
# install/install_macos.sh
# ============================================================
# Installs all dependencies for MinervaDB XL View on macOS
# Supports: macOS 12+ (Intel and Apple Silicon)
#
# Usage:
#   chmod +x install/install_macos.sh
#   ./install/install_macos.sh
# ============================================================

set -euo pipefail

# ---- Colors ----
RED="\033[0;31m"
GREEN="\033[0;32m"
YELLOW="\033[1;33m"
BLUE="\033[0;34m"
BOLD="\033[1m"
NC="\033[0m"  # No Color

# ---- Helper functions ----
info()    { echo -e "${BLUE}[INFO]${NC}  $1"; }
success() { echo -e "${GREEN}[OK]${NC}    $1"; }
warn()    { echo -e "${YELLOW}[WARN]${NC}  $1"; }
error()   { echo -e "${RED}[ERROR]${NC} $1"; exit 1; }
section() { echo -e "\n${BOLD}${BLUE}==>${NC}${BOLD} $1${NC}"; }

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(dirname "$SCRIPT_DIR")"

echo ""
echo -e "${BOLD}${BLUE}"
echo "  ╔══════════════════════════════════════════════╗"
echo "  ║       MinervaDB XL View Installer            ║"
echo "  ║       macOS Setup Script v1.0                ║"
echo "  ╚══════════════════════════════════════════════╝"
echo -e "${NC}"

# ---- Detect architecture ----
ARCH=$(uname -m)
if [ "$ARCH" = "arm64" ]; then
    BREW_PREFIX="/opt/homebrew"
    info "Detected Apple Silicon (M-series) Mac"
else
    BREW_PREFIX="/usr/local"
    info "Detected Intel Mac"
fi

# ============================================================
# Step 1: Check / Install Homebrew
# ============================================================
section "Step 1: Homebrew"
if command -v brew &>/dev/null; then
    success "Homebrew already installed: $(brew --version | head -1)"
else
    info "Installing Homebrew..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    eval "$($BREW_PREFIX/bin/brew shellenv)"
    echo 'eval "$('"$BREW_PREFIX"'/bin/brew shellenv)"' >> ~/.zprofile
    success "Homebrew installed"
fi

# ============================================================
# Step 2: Install Python 3.12
# ============================================================
section "Step 2: Python"
if command -v python3 &>/dev/null; then
    PY_VER=$(python3 --version 2>&1)
    success "Python already installed: $PY_VER"
else
    info "Installing Python 3.12 via Homebrew..."
    brew install python@3.12
    success "Python 3.12 installed"
fi

# Ensure pip is available
if ! command -v pip3 &>/dev/null; then
    warn "pip3 not found, installing..."
    python3 -m ensurepip --upgrade
fi

# ============================================================
# Step 3: Install PostgreSQL Client
# ============================================================
section "Step 3: PostgreSQL Client"
if command -v psql &>/dev/null; then
    success "psql already installed: $(psql --version)"
else
    info "Installing PostgreSQL 16 client tools..."
    brew install postgresql@16
    echo "export PATH=\"$BREW_PREFIX/opt/postgresql@16/bin:\$PATH\"" >> ~/.zshrc
    export PATH="$BREW_PREFIX/opt/postgresql@16/bin:$PATH"
    success "PostgreSQL client installed"
fi

# ============================================================
# Step 4: Create Python Virtual Environment
# ============================================================
section "Step 4: Python Virtual Environment"
cd "$REPO_ROOT"

if [ -d ".venv" ]; then
    warn "Virtual environment already exists, skipping creation"
else
    info "Creating virtual environment..."
    python3 -m venv .venv
    success "Virtual environment created at .venv/"
fi

# Activate venv
source .venv/bin/activate
info "Virtual environment activated"

# ============================================================
# Step 5: Install Python Dependencies
# ============================================================
section "Step 5: Python Dependencies"
info "Upgrading pip..."
pip install --upgrade pip --quiet

info "Installing MinervaDB XL View dependencies..."
pip install -r python/requirements.txt

# Verify key packages
for pkg in psycopg2 pandas openpyxl sqlalchemy; do
    if python3 -c "import $pkg" 2>/dev/null; then
        VER=$(python3 -c "import $pkg; print($pkg.__version__)" 2>/dev/null || echo "installed")
        success "$pkg $VER"
    else
        warn "$pkg not found, trying binary fallback..."
        pip install ${pkg}-binary 2>/dev/null || pip install $pkg
    fi
done

# ============================================================
# Step 6: Configure MinervaDB XL View
# ============================================================
section "Step 6: Configuration"
if [ -f "config/config.ini" ]; then
    warn "config/config.ini already exists, skipping"
else
    cp config/config.example.ini config/config.ini
    success "Created config/config.ini from template"
    warn "Edit config/config.ini with your PostgreSQL credentials before running"
fi

# ============================================================
# Step 7: Create output directory
# ============================================================
section "Step 7: Output Directory"
mkdir -p output
success "Output directory ready at ./output/"

# ============================================================
# Step 8: Test Connection (if config is set)
# ============================================================
section "Step 8: Connection Test"
HOST=$(grep "^host" config/config.ini 2>/dev/null | awk -F= "{print \$2}" | tr -d " " || echo "")

if [ -z "$HOST" ] || [ "$HOST" = "your-postgresql-host" ]; then
    warn "PostgreSQL host not configured. Edit config/config.ini and run:"
    echo "  source .venv/bin/activate"
    echo "  python3 python/pg_connector.py"
else
    info "Testing PostgreSQL connection..."
    if python3 python/pg_connector.py 2>/dev/null; then
        success "Connection test passed!"
    else
        warn "Connection test failed. Check config/config.ini settings."
    fi
fi

# ============================================================
# Install Complete
# ============================================================
echo ""
echo -e "${GREEN}${BOLD}"
echo "  ╔══════════════════════════════════════════════╗"
echo "  ║   MinervaDB XL View Installation Complete!  ║"
echo "  ╚══════════════════════════════════════════════╝"
echo -e "${NC}"

echo ""
echo -e "${BOLD}Next Steps:${NC}"
echo "  1. Edit config/config.ini with your PostgreSQL credentials"
echo "  2. Activate the virtual environment:"
echo "       source .venv/bin/activate"
echo "  3. Test the connection:"
echo "       python3 python/pg_connector.py"
echo "  4. Run the ETL pipeline:"
echo "       python3 python/etl_pipeline.py --pipeline all"
echo "  5. Open the generated Excel file:"
echo "       open output/MinervaDB_XL_View_*.xlsx"
echo ""
echo -e "  📖 Full guide: ${BLUE}docs/MACOS_INSTALL.md${NC}"
echo -e "  ❓ Troubleshooting: ${BLUE}docs/TROUBLESHOOTING.md${NC}"
echo ""
