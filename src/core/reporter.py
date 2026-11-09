import datetime
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

logger = logging.getLogger("InventorAutomation.Reporter")


class RunReporter:
    """Compiles manufacturing run results, DRC status, and exports into Markdown and visual HTML reports."""

    def __init__(self, run_title: str = "CAD Automation & Manufacturing Run Report"):
        self.run_title = run_title
        self.start_time = datetime.datetime.now()
        self.results: List[Dict[str, Any]] = []

    def add_result(
        self,
        part_number: str,
        name: str,
        part_type: str,
        mass_kg: float = 0.0,
        volume_cm3: float = 0.0,
        drc_passed: bool = True,
        drc_details: str = "Passed",
        exports: Optional[List[str]] = None,
    ) -> None:
        """Records the outcome of a single part or assembly automation task."""
        self.results.append(
            {
                "part_number": part_number,
                "name": name,
                "part_type": part_type,
                "mass_kg": mass_kg,
                "volume_cm3": volume_cm3,
                "drc_passed": drc_passed,
                "drc_details": drc_details,
                "exports": exports or [],
            }
        )

    @property
    def total_count(self) -> int:
        return len(self.results)

    @property
    def passed_count(self) -> int:
        return sum(1 for r in self.results if r["drc_passed"])

    @property
    def failed_count(self) -> int:
        return self.total_count - self.passed_count

    def generate_markdown(self, output_path: Optional[Path] = None) -> str:
        """Generates a clean GitHub-flavored Markdown report."""
        timestamp = self.start_time.strftime("%Y-%m-%d %H:%M:%S")
        status_badge = "🟢 SUCCESS" if self.failed_count == 0 else "🔴 ISSUES DETECTED"

        lines = [
            f"# {self.run_title}",
            "",
            f"**Run Status**: {status_badge}  ",
            f"**Execution Timestamp**: `{timestamp}`  ",
            f"**Processed**: {self.total_count} | **Passed**: {self.passed_count} | **Failed**: {self.failed_count}",
            "",
            "## Summary Table",
            "",
            "| Part Number | Name | Type | Mass (kg) | Volume (cm³) | DRC Status | Exports |",
            "| :--- | :--- | :--- | :--- | :--- | :---: | :--- |",
        ]

        for r in self.results:
            drc_icon = "✅ Pass" if r["drc_passed"] else "❌ Fail"
            exports_str = ", ".join(r["exports"]) if r["exports"] else "None"
            lines.append(
                f"| **{r['part_number']}** | {r['name']} | `{r['part_type']}` | "
                f"{r['mass_kg']:.3f} | {r['volume_cm3']:.2f} | {drc_icon} | {exports_str} |"
            )

        lines.extend(
            ["", "---", "*Generated automatically by Autodesk Inventor Automation Engine.*"]
        )
        md_content = "\n".join(lines)

        if output_path:
            p = Path(output_path).expanduser().resolve()
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(md_content, encoding="utf-8")
            logger.info(f"Saved Markdown report to: {p}")

        return md_content

    def generate_html(self, output_path: Optional[Path] = None) -> str:
        """Generates a modern, responsive HTML report with metric cards and styled table."""
        timestamp = self.start_time.strftime("%Y-%m-%d %H:%M:%S")
        status_color = "#10B981" if self.failed_count == 0 else "#EF4444"
        status_text = "ALL CHECKS PASSED" if self.failed_count == 0 else "FAILURES DETECTED"

        rows_html = ""
        for r in self.results:
            badge_class = "badge-pass" if r["drc_passed"] else "badge-fail"
            badge_text = "PASS" if r["drc_passed"] else "FAIL"
            exports_badges = "".join(
                [f'<span class="badge-export">{e.upper()}</span>' for e in r["exports"]]
            )
            rows_html += f"""
            <tr>
                <td><strong>{r['part_number']}</strong></td>
                <td>{r['name']}</td>
                <td><span class="badge-type">{r['part_type']}</span></td>
                <td>{r['mass_kg']:.3f} kg</td>
                <td>{r['volume_cm3']:.2f} cm³</td>
                <td><span class="{badge_class}">{badge_text}</span></td>
                <td>{exports_badges or 'None'}</td>
            </tr>
            """

        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{self.run_title}</title>
    <style>
        :root {{
            --bg: #0F172A;
            --surface: #1E293B;
            --text: #F8FAFC;
            --text-muted: #94A3B8;
            --border: #334155;
            --accent: #3B82F6;
            --success: #10B981;
            --danger: #EF4444;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background: var(--bg);
            color: var(--text);
            margin: 0;
            padding: 40px 20px;
        }}
        .container {{
            max-width: 1100px;
            margin: 0 auto;
        }}
        header {{
            margin-bottom: 30px;
            border-bottom: 1px solid var(--border);
            padding-bottom: 20px;
        }}
        h1 {{
            margin: 0 0 10px 0;
            font-size: 28px;
            font-weight: 700;
        }}
        .timestamp {{
            color: var(--text-muted);
            font-size: 14px;
        }}
        .cards {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .card {{
            background: var(--surface);
            padding: 20px;
            border-radius: 8px;
            border: 1px solid var(--border);
        }}
        .card-label {{
            font-size: 13px;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 6px;
        }}
        .card-value {{
            font-size: 32px;
            font-weight: 700;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            background: var(--surface);
            border-radius: 8px;
            overflow: hidden;
            border: 1px solid var(--border);
        }}
        th, td {{
            padding: 14px 16px;
            text-align: left;
            border-bottom: 1px solid var(--border);
            font-size: 14px;
        }}
        th {{
            background: #243048;
            color: var(--text-muted);
            font-weight: 600;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 0.5px;
        }}
        tr:last-child td {{
            border-bottom: none;
        }}
        .badge-pass {{
            background: rgba(16, 185, 129, 0.2);
            color: var(--success);
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 600;
        }}
        .badge-fail {{
            background: rgba(239, 68, 68, 0.2);
            color: var(--danger);
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 600;
        }}
        .badge-type {{
            background: #334155;
            color: #E2E8F0;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 12px;
            font-family: monospace;
        }}
        .badge-export {{
            background: rgba(59, 130, 246, 0.2);
            color: #60A5FA;
            padding: 2px 6px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
            margin-right: 4px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>{self.run_title}</h1>
            <div class="timestamp">Generated on {timestamp} • Autodesk Inventor Automation Engine</div>
        </header>

        <div class="cards">
            <div class="card">
                <div class="card-label">Overall Status</div>
                <div class="card-value" style="color: {status_color}; font-size: 20px;">{status_text}</div>
            </div>
            <div class="card">
                <div class="card-label">Total Components</div>
                <div class="card-value">{self.total_count}</div>
            </div>
            <div class="card">
                <div class="card-label">Checks Passed</div>
                <div class="card-value" style="color: var(--success);">{self.passed_count}</div>
            </div>
            <div class="card">
                <div class="card-label">Issues / Failures</div>
                <div class="card-value" style="color: var(--danger);">{self.failed_count}</div>
            </div>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Part Number</th>
                    <th>Description</th>
                    <th>Geometry</th>
                    <th>Mass</th>
                    <th>Volume</th>
                    <th>DRC Status</th>
                    <th>Generated Exports</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
    </div>
</body>
</html>
"""
        if output_path:
            p = Path(output_path).expanduser().resolve()
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(html_content, encoding="utf-8")
            logger.info(f"Saved HTML report to: {p}")

        return html_content
