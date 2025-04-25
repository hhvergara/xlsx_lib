curl -LsSf https://astral.sh/uv/install.sh | sh
uv venv
uv pip compile docs/requirements.in \
    --universal \
    --output-file docs/requirements.txt
uv pip sync docs/requirements.txt  --no-cache