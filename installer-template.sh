# Run the installer like this from the Terminal:
# bash installer.sh
# 
# Note: you must include "bash" even if you run this on a different shell, e.g. zsh
# 
# Uploaded to something like s3, you can run it with a single command from the Terminal:
# curl -sSL https://xlwings.s3.amazonaws.com/{{project_placeholder}}/installer.sh | bash

set -e  # stop at errors

MINICONDA_VERSION="Miniconda3-py38_4.8.3-MacOSX-x86_64"
INSTALL_DIR="${HOME}/{{project_placeholder}}"

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
    printf "${YELLOW}Installing packages${NC}\n"
    "$INSTALL_DIR"/bin/conda install appscript=1.1.2 psutil=5.8.0 cryptography=3.4.7 -y
    "$INSTALL_DIR"/bin/pip install --no-deps xlwings==0.23.1
    printf "${YELLOW}Installing xlwings script${NC}\n"
    "$INSTALL_DIR"/bin/xlwings runpython install
    printf "${YELLOW}Copying Data${NC}\n"
    mkdir -p "$INSTALL_DIR"/data
    echo {{version_placeholder}} > "$INSTALL_DIR"/data/version
    printf "${GREEN}Successfully installed {{project_placeholder}}!${NC}\n"
else
    printf "${RED}Please deactivate any conda envs by running 'conda deactivate' before running this command again!${NC}\n"
    exit 1
fi
