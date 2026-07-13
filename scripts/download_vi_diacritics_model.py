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
from huggingface_hub import HfApi, snapshot_download

CACHE_DIR = os.path.join("vi_diacritics_model", "hub_cache")
os.makedirs(CACHE_DIR, exist_ok=True)

REPOS = [
    "vinai/bartpho-syllable",
    "yammdd/vietnamese-error-correction",
]

# v9.4.1: snapshot_download() with no filters grabs EVERY file in the repo,
# including redundant weight formats we never load -- this is what
# ballooned the build from ~388MB to 5.15GB. vinai/bartpho-syllable in
# particular ships BOTH pytorch_model.bin AND tf_model.h5 (TensorFlow) side
# by side for the same BART-LARGE architecture; we only ever load it via
# transformers' PyTorch AutoModelForSeq2SeqLM, so the TF/Flax/ONNX copies
# are pure dead weight. Restrict to what's actually needed:
#   - config/tokenizer files (json, txt, model, vocab)
#   - exactly ONE PyTorch weights format (prefer .safetensors if the repo
#     has it, since some repos ship BOTH .bin and .safetensors for the
#     same weights -- downloading both would silently double the size
#     again even after excluding TF/Flax/ONNX)
IGNORE_PATTERNS_BASE = [
    "*.h5", "tf_model*",           # TensorFlow
    "*.msgpack", "flax_model*",    # Flax/JAX
    "*.onnx", "*.ot",              # ONNX / OpenVINO
    "*.md", "*.png", "*.jpg",      # model cards / images, not needed at runtime
]

api = HfApi()
for repo_id in REPOS:
    files = api.list_repo_files(repo_id)
    has_safetensors = any(f.endswith(".safetensors") for f in files)
    ignore = list(IGNORE_PATTERNS_BASE)
    if has_safetensors:
        # .safetensors present -> skip the redundant .bin copy of the same weights
        ignore.append("*.bin")
    print(f"Downloading {repo_id} into {CACHE_DIR} "
          f"(format: {'safetensors' if has_safetensors else 'pytorch .bin'}) ...")
    snapshot_download(repo_id=repo_id, cache_dir=CACHE_DIR, ignore_patterns=ignore)
    print(f"  done: {repo_id}")

print("All Vietnamese diacritics-restoration model files downloaded (filtered).")
