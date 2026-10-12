# ─────────────────────────────────────────────────────────────
#  Omega 5.0 — Event-Driven PPT Generator Lambda
#  Recommended deployment: AWS Lambda Container Image (ECR)
#  Rationale: opencv + numpy + pillow + lxml exceed the 250 MB
#             unzipped limit for .zip-based Lambda deployments.
# ─────────────────────────────────────────────────────────────
FROM public.ecr.aws/lambda/python:3.13

# System libraries for OpenCV headless + font rendering
RUN dnf install -y \
    mesa-libGL \
    glib2 \
    freetype \
    && dnf clean all

# Install Python dependencies
COPY requirements.txt ${LAMBDA_TASK_ROOT}/
RUN pip install --no-cache-dir -r ${LAMBDA_TASK_ROOT}/requirements.txt

# Copy application code
COPY lambda_function.py ${LAMBDA_TASK_ROOT}/

# Lambda handler entrypoint
CMD ["lambda_function.lambda_handler"]
