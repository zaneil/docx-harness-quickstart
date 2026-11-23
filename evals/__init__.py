"""Evals: tasks a real model is asked to do with the docx-harness tools, each
with automatic checks on the document it produces.

    uv run python -m evals list
    uv run python -m evals run --model openai/gpt-6-luna

A task is one YAML file under ``evals/tasks/<category>/``. See ``evals/README.md``.
"""
