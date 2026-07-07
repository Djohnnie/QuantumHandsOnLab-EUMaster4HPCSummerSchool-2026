@echo off
REM Convert the slide-NN.svg files in _slides\ into slides.pptx.
setlocal
python "%~dp0_tools\svg_to_pptx.py" %*
endlocal
