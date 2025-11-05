# Technology Stack - Wan2.2 Animate

## System Specifications
- **Platform**: Windows (win32)
- **Shell**: CMD/PowerShell
- **Hardware**: ASUS ROG Zephyrus M16
  - CPU: Intel Core i9-12900H
  - RAM: 32GB
  - GPU: NVIDIA RTX 3080 Ti 16GB VRAM
  - Storage: SSD NVMe

## Current Environment Status
- **Python**: 3.13.3 (needs downgrade to 3.10 for compatibility)
- **CUDA**: 12.5 (compatible)
- **Conda**: 24.9.2 (installed but not in PATH)
- **PyTorch**: Not installed (needs torch>=2.4.0)
- **Existing Conda Envs**: ltxvideo, trellis, uvixComfy, asset_extractor, etc.

## Required Tech Stack
- **Python**: 3.10 (recommended for AI/ML compatibility)
- **PyTorch**: 2.4.1 with CUDA 12.1 support
- **CUDA Toolkit**: 12.1 or 12.5 (already available)
- **Package Manager**: Conda (isolated environment)
- **Web Interface**: Gradio for user interaction

## Key Dependencies
```txt
# Core ML Framework
torch>=2.4.0
torchvision>=0.19.0
diffusers>=0.31.0
transformers>=4.49.0,<=4.51.3
accelerate>=1.1.1

# Animate-specific
decord
peft
onnxruntime
SAM-2 (from git)

# Interface & Utils
gradio (for web interface)
opencv-python>=4.9.0.80
imageio[ffmpeg]
flash_attn (may need special installation)
```

## Model Requirements
- **Wan2.2-Animate-14B**: ~27GB model weights
- **Process Checkpoint**: Additional preprocessing models
- **Storage**: ~50GB total for all model components

## Performance Expectations
- **720P Generation**: Supported with memory optimization
- **1080P Generation**: Potentially supported (needs testing)
- **Memory Usage**: 16GB VRAM with offloading techniques
- **Processing Time**: Several minutes per video depending on length

## Wan-Specific Optimization Flags
```cmd
# Memory optimization for RTX 3080 Ti (16GB VRAM)
--offload_model True          # Offload model parts to CPU when not in use
--convert_model_dtype         # Convert to optimal data types
--t5_cpu                      # Run T5 text encoder on CPU to save VRAM

# Performance tuning
--ulysses_size 8             # For multi-GPU setups (future expansion)
--dit_fsdp --t5_fsdp         # Distributed training flags (multi-GPU)
```

## GPU Memory Monitoring
- Monitor VRAM usage with `nvidia-smi`
- Expected usage: ~14-16GB for 720P generation
- If OOM occurs: reduce resolution or enable more optimization flags
- Preprocessing requires additional ~2-4GB temporarily

## Installation Commands
```cmd
# Fix Conda PATH
set PATH=%USERPROFILE%\Anaconda3\Scripts;%USERPROFILE%\Anaconda3;%PATH%

# Create isolated environment
conda create -n wan_animate python=3.10 -y
conda activate wan_animate

# Install PyTorch with CUDA support
pip install torch==2.4.1 torchvision==0.19.1 --index-url https://download.pytorch.org/whl/cu121

# Install project dependencies
pip install -r requirements.txt
pip install -r requirements_animate.txt

# Install Gradio for interface
pip install gradio
```