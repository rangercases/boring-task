import os
import sys

# Ensure current directory is on sys.path
app_dir = os.path.dirname(os.path.abspath(__file__))
if app_dir not in sys.path:
    sys.path.insert(0, app_dir)

from claim_helper import ClaimHelperAppleApp

if __name__ == "__main__":
    app = ClaimHelperAppleApp()
    app.mainloop()
