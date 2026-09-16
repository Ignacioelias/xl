"""tcfeed: declarative think-cell feed blocks. See references/pipeline.md.

  py -3.14 tcfeed_cli.py plan    SPEC.json                 dry run (nothing written)
  py -3.14 tcfeed_cli.py run     SPEC.json [--engine-check] build -> compute -> inject -> calcPr -> verify
  py -3.14 tcfeed_cli.py compute SPEC.json [-o values.json] values for `xl inject --from`
  py -3.14 tcfeed_cli.py verify  RUN_DIR [--target FILE]    re-run the verification
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from tcfeed.pipeline import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
