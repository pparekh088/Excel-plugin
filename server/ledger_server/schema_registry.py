"""Registry of the typed tool surface (INV-1 enforcement point).

The Zod definitions in addin/src/tools/schemas.ts are the source of truth.
`npm run schemas` in addin/ exports them as JSON Schema into shared/schemas/,
which is checked in so the server never needs a Node toolchain. This module
loads those schemas at startup and validates every tool call's params and
result against them. Anything that fails validation is rejected before it can
reach a workbook.

Files per tool (dots in tool names are kept as-is):
    <tool>.params.json   e.g. range.read.params.json
    <tool>.result.json   e.g. range.read.result.json
    manifest.json        tool metadata: access level, risk tier, description
"""

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from jsonschema import Draft202012Validator

Access = Literal["read", "write", "control"]
# Risk tiers per the handoff (section 4). Read/inspection tools carry "none".
Risk = Literal["none", "low", "medium", "high"]


@dataclass(frozen=True)
class ToolSpec:
    name: str
    access: Access
    risk: Risk
    description: str
    params_validator: Draft202012Validator
    result_validator: Draft202012Validator


class SchemaValidationError(Exception):
    """A tool call's params or result failed JSON Schema validation."""

    def __init__(self, kind: str, tool: str, errors: list[dict[str, Any]]):
        self.kind = kind
        self.tool = tool
        self.errors = errors
        super().__init__(f"{kind} validation failed for tool '{tool}'")


class ToolRegistry:
    def __init__(self, specs: dict[str, ToolSpec]):
        self._specs = specs

    @classmethod
    def load(cls, schemas_dir: Path) -> "ToolRegistry":
        manifest_path = schemas_dir / "manifest.json"
        if not manifest_path.exists():
            raise FileNotFoundError(
                f"Tool schema manifest not found at {manifest_path}. "
                "Run `npm run schemas` in addin/ to generate shared/schemas."
            )
        manifest = json.loads(manifest_path.read_text())
        specs: dict[str, ToolSpec] = {}
        for name, meta in manifest["tools"].items():
            params_schema = json.loads((schemas_dir / f"{name}.params.json").read_text())
            result_schema = json.loads((schemas_dir / f"{name}.result.json").read_text())
            Draft202012Validator.check_schema(params_schema)
            Draft202012Validator.check_schema(result_schema)
            specs[name] = ToolSpec(
                name=name,
                access=meta["access"],
                risk=meta["risk"],
                description=meta.get("description", ""),
                params_validator=Draft202012Validator(params_schema),
                result_validator=Draft202012Validator(result_schema),
            )
        return cls(specs)

    @property
    def tool_names(self) -> list[str]:
        return sorted(self._specs)

    def get(self, tool: str) -> ToolSpec | None:
        return self._specs.get(tool)

    @staticmethod
    def _collect_errors(validator: Draft202012Validator, instance: Any) -> list[dict[str, Any]]:
        return [
            {
                "path": "/" + "/".join(str(p) for p in err.absolute_path),
                "message": err.message,
            }
            for err in sorted(validator.iter_errors(instance), key=lambda e: list(e.absolute_path))
        ]

    def validate_params(self, tool: str, params: Any) -> None:
        spec = self._specs[tool]
        errors = self._collect_errors(spec.params_validator, params)
        if errors:
            raise SchemaValidationError("params", tool, errors)

    def validate_result(self, tool: str, result: Any) -> None:
        spec = self._specs[tool]
        errors = self._collect_errors(spec.result_validator, result)
        if errors:
            raise SchemaValidationError("result", tool, errors)
