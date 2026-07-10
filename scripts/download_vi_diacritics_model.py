"""
scripts/download_vi_diacritics_model.py

Run this BEFORE pyinstaller, during the build (see .github/workflows/build.yml).
NOT run on the end user's machine -- this bakes the model into the build
output so end users never need internet access for it.

Downloads both:
  - vinai/bartpho-syllable          (base model)
  - yammdd/vietnamese-error-correction  (LoRA adapter on top of the base model)

into a real HuggingFace hub cache folder structure (models--org--name/...)
under vi_diacritics_model/hub_cache/. That folder gets bundled into
_internal at build time via the `datas` entry in SmartSearchAI.spec, and at
runtime app.py points HF_HOME at that exact bundled folder before importing
transformers anywhere, so both the adapter AND its base model resolve
completely offline via local_files_only=True -- no network call needed on
whatever machine actually runs the built exe.

Both repos must be downloaded into the SAME cache_dir (not two separate
folders) because transformers' automatic PEFT support resolves the base
model against the same HF cache the adapter itself was loaded from.
"""

import os
from huggingface_hub import snapshot_download

CACHE_DIR = os.path.join("vi_diacritics_model", "hub_cache")
os.makedirs(CACHE_DIR, exist_ok=True)

REPOS = [
    "vinai/bartpho-syllable",
    "yammdd/vietnamese-error-correction",
]

for repo_id in REPOS:
    print(f"Downloading {repo_id} into {CACHE_DIR} ...")
    snapshot_download(repo_id=repo_id, cache_dir=CACHE_DIR)
    print(f"  done: {repo_id}")

print("All Vietnamese diacritics-restoration model files downloaded.")
