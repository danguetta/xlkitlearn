# Run installer like this:
# bash installer.sh
# 
# Uploaded to something like s3, you can run it like so:
# curl -sSL https://xlwings.s3.amazonaws.com/xlkitlearn/installer.sh | bash
#
# In case of permission errors use 'sudo bash installer.sh' or if hosted:
# curl -sSL https://xlwings.s3.amazonaws.com/xlkitlearn/installer.sh | sudo bash

set -e  # stop at errors

MINICONDA_VERSION="Miniconda3-py38_22.11.1-1-MacOSX-x86_64"
INSTALL_DIR="${HOME}/xlkitlearn"

GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[0;33m'
NC='\033[0m' # No Color

if [ -z "$CONDA_DEFAULT_ENV" ];then
    # We're not in an activated conda environment and installation can start
    printf "${YELLOW}Cleaning up existing installation${NC}\n"
    rm -rf "$INSTALL_DIR" || true
    printf "${YELLOW}Downloading Miniconda${NC}\n"
    curl -L https://repo.anaconda.com/miniconda/"$MINICONDA_VERSION".sh -o /tmp/"$MINICONDA_VERSION".sh
    printf "${YELLOW}Installing Miniconda${NC}\n"
    bash /tmp/"$MINICONDA_VERSION".sh -u -b -p "$INSTALL_DIR"
    cat > "$INSTALL_DIR"/requirements.txt << 'EOF'
{{requirements_placeholder}}
EOF

    printf "${YELLOW}Installing packages${NC}\n"
    "$INSTALL_DIR"/bin/pip install -r "$INSTALL_DIR"/requirements.txt
    printf "${YELLOW}Installing xlwings script${NC}\n"
    "$INSTALL_DIR"/bin/xlwings runpython install
    printf "${YELLOW}Copying Data${NC}\n"
    mkdir -p "$INSTALL_DIR"/data
    echo {{version_placeholder}} > "$INSTALL_DIR"/data/version
    printf "${YELLOW}Compiling packages${NC}\n"
    "$INSTALL_DIR"/bin/python -c "python_command_placeholder"
    printf "${YELLOW}Downloading add-in Excel${NC}\n"
    curl -L -o ~/Desktop/XLKitLearn.xltm "https://github.com/danguetta/XLKitLearn/releases/latest/download/XLKitLearn.xltm"
    printf "${GREEN}Successfully installed XLKitLearn!${NC}\n"
else
    printf "${RED}I need a little help to complete the installation. Please type the words 'conda activate' (without the quotes) below and press enter. Then, re-run exactly the same command you just ran.${NC}\n"
    exit 1
fi
