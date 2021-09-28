Repository to ease the hard work of an evalens deleguate

## To use the code
Place all your files in the input folder and run:

    python -m main

or to set some parameters:

    python -m main with input_folder=<INPUT_PATH> output_folder=<OUTPUT_PATH> output_prefix=<PREFIX>

where:
* ***INTPUT_PATH*** is the path to the folder containing the files to be processed
  
  Default value: "input"
* ***OUTPUT_PATH*** is the path where you want to output all the processed files
  
  Default value: "output"
* ***PREFIX*** is the prefix you want to add to all the processed files
  
  Default value: "[PROCESSED] "
