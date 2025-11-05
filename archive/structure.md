# Project Structure - Wan2.2 Animate

## Current Organization

```
.
├── .kiro/                   # Kiro AI assistant configuration
│   ├── steering/           # AI guidance documents
│   │   ├── product.md      # Wan2.2 Animate project overview
│   │   ├── tech.md         # Hardware specs and tech stack
│   │   └── structure.md    # This file - project organization
│   └── specs/              # Feature specifications and tasks
├── .vscode/                # VSCode configuration
│   └── settings.json       # Editor settings
├── Wan2.2-main/            # Downloaded Wan2.2 repository
│   └── Wan2.2-main/        # Actual project files
│       ├── wan/            # Core model modules
│       ├── examples/       # Sample inputs and demos
│       ├── generate.py     # Main inference script
│       ├── requirements.txt # Base dependencies
│       └── requirements_animate.txt # Animate-specific deps
└── models/                 # Model weights storage (to be created)
    └── Wan2.2-Animate-14B/ # Downloaded model weights
```

## Planned Structure (Post-Installation)

```
.
├── .kiro/                  # Kiro configuration
├── Wan2.2-main/           # Original repository
├── models/                # Model weights and checkpoints
│   ├── Wan2.2-Animate-14B/
│   └── process_checkpoint/
├── interface/             # Custom Gradio interface
│   ├── app.py            # Main Gradio application
│   ├── utils.py          # Helper functions
│   └── templates/        # UI templates
├── workspace/            # User workspace for processing
│   ├── inputs/           # User uploaded videos/images
│   ├── outputs/          # Generated results
│   └── temp/             # Temporary processing files
└── conda_env/            # Environment setup scripts
    └── wan_animate.yml   # Conda environment definition
```

## File Organization Conventions

### Model Storage
- All model weights in `models/` directory
- Separate subdirectories for each model variant
- Keep original directory structure from downloads

### User Interface
- Gradio app in dedicated `interface/` directory
- Modular design with separate utility functions
- Template-based UI for consistency

### Processing Workflow
- Input files in `workspace/inputs/`
- Preprocessing results in `workspace/temp/`
- Final outputs in `workspace/outputs/`
- Automatic cleanup of temporary files

### Environment Management
- Dedicated conda environment: `wan_animate`
- Environment definition files for reproducibility
- Separate requirements for different model components

## Development Phases

### Phase 1: Basic Setup
- [x] Download repository
- [ ] Create conda environment
- [ ] Install dependencies
- [ ] Download model weights

### Phase 2: Core Functionality
- [ ] Test basic inference
- [ ] Implement preprocessing pipeline
- [ ] Verify animation and replacement modes

### Phase 3: User Interface
- [ ] Create Gradio interface
- [ ] Implement file upload/download
- [ ] Add progress tracking
- [ ] Error handling and validation

### Phase 4: Optimization
- [ ] Memory optimization for 16GB VRAM
- [ ] Performance tuning
- [ ] Resolution testing (720P/1080P)
- [ ] Batch processing capabilities