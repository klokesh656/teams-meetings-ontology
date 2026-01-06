"""
Setup Checker - Verify Prerequisites
=====================================
Checks if all required tools and services are configured correctly.
Run this before attempting transcription.
"""

import os
import sys
import subprocess
from pathlib import Path
from dotenv import load_dotenv

# Colors for output (Windows compatible)
GREEN = "\033[92m"
RED = "\033[91m"
YELLOW = "\033[93m"
RESET = "\033[0m"

def print_status(check_name, status, message=""):
    """Print check status"""
    if status:
        symbol = "[OK]"
        color = GREEN
    else:
        symbol = "[FAIL]"
        color = RED
    
    print(f"{color}{symbol}{RESET} {check_name}")
    if message:
        print(f"      {message}")

def check_ffmpeg():
    """Check if FFmpeg is installed"""
    try:
        result = subprocess.run(['ffmpeg', '-version'], capture_output=True, text=True)
        if result.returncode == 0:
            version = result.stdout.split('\n')[0]
            return True, version
        return False, "FFmpeg command failed"
    except FileNotFoundError:
        return False, "FFmpeg not found in PATH"

def check_python_packages():
    """Check if required Python packages are installed"""
    required = [
        'azure.identity',
        'azure.storage.blob',
        'azure.cognitiveservices.speech',
        'requests',
        'pandas',
        'openpyxl',
        'openai',
        'dotenv'
    ]
    
    missing = []
    for package in required:
        try:
            __import__(package.replace('-', '_'))
        except ImportError:
            missing.append(package)
    
    if missing:
        return False, f"Missing: {', '.join(missing)}"
    return True, f"All {len(required)} packages installed"

def check_env_variables():
    """Check if required environment variables are set"""
    load_dotenv()
    
    required = {
        'AZURE_CLIENT_ID': 'Azure AD Client ID',
        'AZURE_CLIENT_SECRET': 'Azure AD Client Secret',
        'AZURE_TENANT_ID': 'Azure Tenant ID',
        'AZURE_STORAGE_CONNECTION_STRING': 'Azure Blob Storage',
        'AZURE_OPENAI_ENDPOINT': 'Azure OpenAI Endpoint',
        'AZURE_OPENAI_KEY': 'Azure OpenAI Key',
        'AZURE_SPEECH_KEY': 'Azure Speech Service Key (NEW!)',
        'AZURE_SPEECH_REGION': 'Azure Speech Region'
    }
    
    missing = []
    configured = []
    
    for key, name in required.items():
        value = os.getenv(key)
        if not value or value == 'YOUR_SPEECH_SERVICE_KEY_HERE':
            missing.append(name)
        else:
            configured.append(name)
    
    if missing:
        return False, f"Missing: {', '.join(missing)}"
    return True, f"All {len(configured)} variables configured"

def check_directories():
    """Check if required directories exist"""
    required = ['transcripts', 'recordings', 'output', 'logs']
    
    for dir_name in required:
        Path(dir_name).mkdir(exist_ok=True)
    
    return True, f"All directories ready: {', '.join(required)}"

def check_azure_speech_sdk():
    """Check if Azure Speech SDK is properly installed"""
    try:
        import azure.cognitiveservices.speech as speechsdk
        return True, f"Azure Speech SDK version: {speechsdk.__version__}"
    except ImportError:
        return False, "Install: pip install azure-cognitiveservices-speech"
    except Exception as e:
        return False, str(e)

def main():
    """Run all checks"""
    print("\n" + "="*60)
    print("TRANSCRIPTION SETUP CHECKER")
    print("="*60 + "\n")
    
    all_passed = True
    
    # Check 1: FFmpeg
    print("1. Checking FFmpeg...")
    passed, message = check_ffmpeg()
    print_status("FFmpeg", passed, message)
    if not passed:
        print(f"{YELLOW}      → Install: choco install ffmpeg{RESET}")
    all_passed = all_passed and passed
    print()
    
    # Check 2: Python Packages
    print("2. Checking Python Packages...")
    passed, message = check_python_packages()
    print_status("Python Packages", passed, message)
    if not passed:
        print(f"{YELLOW}      → Install: pip install -r requirements.txt{RESET}")
    all_passed = all_passed and passed
    print()
    
    # Check 3: Azure Speech SDK
    print("3. Checking Azure Speech SDK...")
    passed, message = check_azure_speech_sdk()
    print_status("Azure Speech SDK", passed, message)
    if not passed:
        print(f"{YELLOW}      → Install: pip install azure-cognitiveservices-speech{RESET}")
    all_passed = all_passed and passed
    print()
    
    # Check 4: Environment Variables
    print("4. Checking Environment Variables...")
    passed, message = check_env_variables()
    print_status("Environment Variables", passed, message)
    if not passed:
        print(f"{YELLOW}      → Configure .env file (see TRANSCRIPTION_QUICKSTART.md){RESET}")
    all_passed = all_passed and passed
    print()
    
    # Check 5: Directories
    print("5. Checking Directories...")
    passed, message = check_directories()
    print_status("Directories", passed, message)
    print()
    
    # Summary
    print("="*60)
    if all_passed:
        print(f"{GREEN}✓ ALL CHECKS PASSED! Ready to transcribe!{RESET}")
        print("\nNext steps:")
        print("  1. Run: python src/transcribe_recordings.py")
        print("  2. Check results: dir transcripts\\*transcribed.vtt")
    else:
        print(f"{RED}✗ Some checks failed. Please fix the issues above.{RESET}")
        print("\nQuick fixes:")
        print("  • FFmpeg: choco install ffmpeg -y")
        print("  • Packages: pip install -r requirements.txt")
        print("  • .env: Add AZURE_SPEECH_KEY and AZURE_SPEECH_REGION")
        print("\nSee TRANSCRIPTION_QUICKSTART.md for detailed setup instructions.")
    print("="*60 + "\n")
    
    return 0 if all_passed else 1

if __name__ == '__main__':
    sys.exit(main())
