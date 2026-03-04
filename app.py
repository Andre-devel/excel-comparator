"""
PONTO DE ENTRADA DA APLICAÇÃO
Servidor Flask que expõe a interface web e a API de comparação.
"""

import os
import uuid
from pathlib import Path

from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

from reader import read_excel, FileReadError
from validator import validate_structure
from comparator import compare
from reporter import generate_report, generate_merge_report

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50 MB

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


# ---------------------------------------------------------------------------
# Utilitários
# ---------------------------------------------------------------------------

def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def save_upload(file, label: str) -> Path:
    """Salva o arquivo enviado e retorna o caminho."""
    if not file or file.filename == "":
        raise ValueError(f"{label}: nenhum arquivo enviado.")
    if not allowed_file(file.filename):
        raise ValueError(
            f"{label}: formato inválido. Apenas arquivos .xlsx são aceitos."
        )
    filename = f"{uuid.uuid4()}_{secure_filename(file.filename)}"
    path = UPLOAD_DIR / filename
    file.save(path)
    return path


def parse_list_param(value: str) -> list[str]:
    """Converte 'col1, col2, col3' em ['col1', 'col2', 'col3']."""
    if not value or not value.strip():
        return []
    return [c.strip() for c in value.split(",") if c.strip()]


# ---------------------------------------------------------------------------
# Rotas
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/compare", methods=["POST"])
def api_compare():
    path1 = path2 = output_path = None

    try:
        # 1. Receber e salvar arquivos
        file1 = request.files.get("file1")
        file2 = request.files.get("file2")

        file1_name = file1.filename if file1 else "Arquivo 1"
        file2_name = file2.filename if file2 else "Arquivo 2"

        try:
            path1 = save_upload(file1, "Arquivo 1")
            path2 = save_upload(file2, "Arquivo 2")
        except ValueError as e:
            return jsonify({"success": False, "error": str(e)}), 400

        # 2. Parâmetros opcionais
        primary_key = request.form.get("primary_key", "").strip() or None
        ignore_columns = parse_list_param(request.form.get("ignore_columns", ""))

        # 3. Leitura dos arquivos
        try:
            df1 = read_excel(str(path1), "Arquivo 1")
        except FileReadError as e:
            return jsonify({"success": False, "error": str(e)}), 422

        try:
            df2 = read_excel(str(path2), "Arquivo 2")
        except FileReadError as e:
            return jsonify({"success": False, "error": str(e)}), 422

        # 4. Validação estrutural
        validation = validate_structure(df1, df2, primary_key, ignore_columns)
        if not validation.valid:
            error_msg = "Validação estrutural falhou:\n• " + "\n• ".join(validation.errors)
            return jsonify({"success": False, "error": error_msg}), 422

        # 5. Comparação
        result = compare(df1, df2, primary_key, ignore_columns)

        # 6. Geração do relatório
        output_filename = f"comparacao_{uuid.uuid4()}.xlsx"
        output_path = OUTPUT_DIR / output_filename
        generate_report(df1, result, str(output_path), file1_name, file2_name)

        # 7. Resposta de sucesso
        return jsonify({
            "success": True,
            "stats": {
                "total_rows": result.total_rows_compared,
                "total_divergences": result.total_divergences,
                "divergent_rows": len(result.divergent_rows),
            },
            "download_url": f"/api/download/{output_filename}",
            "divergences": [
                {
                    "row": d.row_number,
                    "column": d.column_name,
                    "value1": d.value_file1,
                    "value2": d.value_file2,
                }
                for d in result.divergences[:200]  # Preview: até 200 linhas na UI
            ],
        })

    except Exception as e:
        app.logger.exception("Erro inesperado durante comparação")
        return jsonify({
            "success": False,
            "error": f"Erro interno inesperado: {str(e)}"
        }), 500

    finally:
        # Limpar arquivos temporários de upload
        for path in [path1, path2]:
            if path and Path(path).exists():
                try:
                    Path(path).unlink()
                except Exception:
                    pass


@app.route("/api/merge", methods=["POST"])
def api_merge():
    path1 = path2 = None
    try:
        file1 = request.files.get("file1")
        file2 = request.files.get("file2")

        file1_name = file1.filename if file1 else "Arquivo 1"
        file2_name = file2.filename if file2 else "Arquivo 2"

        try:
            path1 = save_upload(file1, "Arquivo 1")
            path2 = save_upload(file2, "Arquivo 2")
        except ValueError as e:
            return jsonify({"success": False, "error": str(e)}), 400

        primary_key = request.form.get("primary_key", "").strip() or None
        ignore_columns = parse_list_param(request.form.get("ignore_columns", ""))

        try:
            df1 = read_excel(str(path1), "Arquivo 1")
        except FileReadError as e:
            return jsonify({"success": False, "error": str(e)}), 422

        try:
            df2 = read_excel(str(path2), "Arquivo 2")
        except FileReadError as e:
            return jsonify({"success": False, "error": str(e)}), 422

        # Valida estrutura, mas ignora divergência de número de linhas (permitida na junção)
        validation = validate_structure(df1, df2, primary_key, ignore_columns)
        if not validation.valid:
            structural_errors = [
                e for e in validation.errors
                if "número de linhas" not in e.lower()
            ]
            if structural_errors:
                return jsonify({
                    "success": False,
                    "error": "Validação falhou:\n• " + "\n• ".join(structural_errors)
                }), 422

        output_filename = f"conferencia_{uuid.uuid4()}.xlsx"
        output_path = OUTPUT_DIR / output_filename
        generate_merge_report(df1, df2, str(output_path), file1_name, file2_name, primary_key, ignore_columns)

        return jsonify({
            "success": True,
            "download_url": f"/api/download/{output_filename}",
        })

    except Exception as e:
        app.logger.exception("Erro inesperado durante junção")
        return jsonify({"success": False, "error": f"Erro interno inesperado: {str(e)}"}), 500

    finally:
        for path in [path1, path2]:
            if path and Path(path).exists():
                try:
                    Path(path).unlink()
                except Exception:
                    pass


@app.route("/api/download/<filename>")
def download(filename):
    """Download do arquivo de relatório gerado."""
    # Segurança: impede path traversal
    safe_name = Path(secure_filename(filename)).name
    file_path = OUTPUT_DIR / safe_name

    if not file_path.exists():
        return jsonify({"error": "Arquivo não encontrado ou expirado."}), 404

    return send_file(
        str(file_path),
        as_attachment=True,
        download_name="relatorio_comparacao.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Inicialização
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
