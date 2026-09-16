"""tcfeed: declarative think-cell feed blocks for an existing Excel model.

spec (JSON) -> layout plan -> compute (from cached values) -> COM build (no calculation)
-> xl inject -> xl calcpr restore -> verify.

Entry point: scripts/tcfeed_cli.py (or `python -m tcfeed` with scripts/ on sys.path).
Reference: references/pipeline.md.
"""
__version__ = "1.0.0"
