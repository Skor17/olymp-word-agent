@echo off
rem ==========================================================================
rem  word-agent demo — double-click this file.
rem  It: (1) makes sure a virtualenv exists, (2) regenerates the sample
rem  document, (3) runs the agent on the sample instruction, and shows the
rem  integrity-verification report. No API key is needed (mock planner).
rem ==========================================================================
setlocal
cd /d "%~dp0"

echo.
echo === word-agent demo ===
echo Working directory: %cd%
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [setup] No virtualenv found - creating one and installing dependencies...
    python -m venv .venv
    if errorlevel 1 (
        echo [setup] FAILED to create the virtualenv. Is Python on PATH?
        pause
        exit /b 1
    )
    call ".venv\Scripts\python.exe" -m pip install --upgrade pip -q
    call ".venv\Scripts\python.exe" -m pip install -r requirements.txt -q
    if errorlevel 1 (
        echo [setup] FAILED to install dependencies.
        pause
        exit /b 1
    )
)

echo [1/3] Generating sample input document...
call ".venv\Scripts\python.exe" examples\make_sample_doc.py
if errorlevel 1 goto :fail

echo.
echo [2/3] Running the agent (mock planner, offline)...
call ".venv\Scripts\python.exe" -m word_agent.cli --planner mock examples\input.docx -o examples\output.docx --instruction-file examples\instruction.txt
if errorlevel 1 goto :fail

echo.
echo [3/3] Running the test suite...
call ".venv\Scripts\python.exe" -m pytest tests -q
if errorlevel 1 goto :fail

echo.
echo ============================================================
echo  Done. Compare examples\input.docx and examples\output.docx
echo  in Word - the three planned changes are applied and all other
echo  content, formatting, the image, and existing rows are intact.
echo ============================================================
pause
exit /b 0

:fail
echo.
echo DEMO FAILED - see the error output above.
pause
exit /b 1
