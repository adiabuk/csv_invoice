#!/usr/bin/env bash

# Install requirements

apt update
apt install python3 python3-pip pandoc texlive-latex-base texlive-fonts-recommended texlive-latex-extra
locale-gen en_GB.UTF-8
update-local
