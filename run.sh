VENV="./venv"
if ! [ -d "$VENV" ]; then
  pip install virtualenv
  virtualenv -p python3 $VENV || virtualenv -p python $VENV
  source $VENV/bin/activate
  $VENV/bin/pip install -r requirements.txt
else
  source $VENV/bin/activate
fi

python3 -m main || python -m main
