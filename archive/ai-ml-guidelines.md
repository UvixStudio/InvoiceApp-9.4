# AI/ML Development Guidelines

## Model Management
- Always validate model inputs and outputs
- Implement proper error handling for model loading failures
- Use version control for model weights and configurations
- Test with different input sizes and formats
- Implement fallback mechanisms for model failures

## GPU Memory Management
- Monitor GPU memory usage during development and runtime
- Implement memory optimization techniques:
  - Model offloading (`--offload_model`)
  - Data type conversion (`--convert_model_dtype`)
  - CPU fallback for large components (`--t5_cpu`)
- Use memory profiling tools to identify bottlenecks
- Implement graceful degradation when memory is insufficient

## Performance Optimization
- Profile inference time and memory usage
- Implement batch processing where applicable
- Use appropriate precision (FP16/FP32) based on hardware
- Consider model quantization for deployment
- Monitor and log performance metrics

## Data Handling
- Validate input data formats and sizes
- Implement proper preprocessing pipelines
- Handle edge cases (corrupted files, unsupported formats)
- Use secure temporary file handling
- Clean up temporary files after processing

## Error Handling & Recovery
- Implement comprehensive error handling for:
  - CUDA out of memory errors
  - Model loading failures
  - Invalid input data
  - Network/download issues
- Provide clear error messages with suggested solutions
- Implement retry mechanisms where appropriate
- Log errors for debugging and monitoring

## Testing & Validation
- Test with various input sizes and formats
- Validate model outputs for correctness
- Implement unit tests for preprocessing functions
- Test memory optimization flags
- Validate GPU compatibility and performance

## Deployment Considerations
- Document hardware requirements clearly
- Provide installation scripts and environment setup
- Implement health checks for model availability
- Monitor resource usage in production
- Plan for model updates and versioning