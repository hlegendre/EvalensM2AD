VENV="./venv"
if ! [ -d "$VENV" ]; then
  sudo apt install python3.9-venv
  python -m venv $VENV || python3 -m venv $VENV
  source $VENV/bin/activate
  $VENV/bin/pip install -r requirements.txt
else
  source $VENV/bin/activate
fi

python -m main || python3 -m main
