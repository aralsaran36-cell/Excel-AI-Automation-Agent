
import runpy
import sys
from pathlib import Path

# Project root directory
BASE_DIR = Path(__file__).resolve().parent

# Add project root to Python path
sys.path.insert(0, str(BASE_DIR))


def main():

    print("=" * 50)
    print("       EXCEL AI AGENT")
    print("=" * 50)

    try:
        print("\nStarting Backend Agent...\n")

        # Launch existing Backend Agent
        runpy.run_module(
            "backend.agent",
            run_name="__main__"
        )

    except KeyboardInterrupt:
        print("\nExcel AI Agent stopped by user.")

    except Exception as e:
        print(f"\nApplication Error: {e}")


if __name__ == "__main__":
    main()
