from flask import Flask, render_template, request, redirect, url_for, send_file
import sqlite3
from datetime import datetime
import pandas as pd
import io
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

app = Flask(__name__)
DB = "database.db"


# Página inicial
@app.route("/")
def index():
    return render_template("index.html")


# Cadastro de unidades, áreas e temas
@app.route("/cadastrar", methods=["GET", "POST"])
def cadastrar():
    msg = ""
    with sqlite3.connect(DB) as conn:
        cur = conn.cursor()

        if request.method == "POST":
            tipo = request.form.get("tipo")
            modo = request.form.get("modo")

            if modo == "simples":
                nome = request.form.get("nome")
                if tipo == "unidade":
                    sigla = request.form.get("sigla")
                    cur.execute("INSERT INTO unidades (nome, sigla) VALUES (?, ?)", (nome, sigla))
                    msg = "Unidade cadastrada com sucesso."
                elif tipo == "area":
                    cur.execute("INSERT INTO areas (nome) VALUES (?)", (nome,))
                    msg = "Área cadastrada com sucesso."
                elif tipo == "tema":
                    cur.execute("INSERT INTO temas (nome) VALUES (?)", (nome,))
                    msg = "Tema cadastrada com sucesso."

            elif modo == "lote":
                linhas = request.form.get("lote", "").split("\n")
                for linha in linhas:
                    nome = linha.strip()
                    if not nome:
                        continue
                    if tipo == "unidade":
                        if " - " in nome:
                            nome_unidade, sigla = nome.split(" - ", 1)
                            cur.execute("INSERT INTO unidades (nome, sigla) VALUES (?, ?)", (nome_unidade.strip(), sigla.strip()))
                    elif tipo == "area":
                        cur.execute("INSERT INTO areas (nome) VALUES (?)", (nome,))
                    elif tipo == "tema":
                        cur.execute("INSERT INTO temas (nome) VALUES (?)", (nome,))
                msg = "Cadastro em lote realizado com sucesso."

        unidades = cur.execute("SELECT * FROM unidades").fetchall()
        areas = cur.execute("SELECT * FROM areas").fetchall()
        temas = cur.execute("SELECT * FROM temas").fetchall()

    return render_template("cadastrar.html", msg=msg, unidades=unidades, areas=areas, temas=temas)


# Criação e visualização de cronogramas
@app.route("/cronograma", methods=["GET", "POST"])
def cronograma():
    msg = ""
    with sqlite3.connect(DB) as conn:
        cur = conn.cursor()

        if request.method == "POST":
            tema = request.form.get("tema")
            data = request.form.get("data")
            hora = request.form.get("hora")
            formato = request.form.get("formato")
            unidades = request.form.getlist("unidades")
            areas = request.form.getlist("areas")

            unidades_str = ", ".join(unidades)
            areas_str = ", ".join(areas)

            cur.execute("""
                INSERT INTO cronogramas (tema_id, data, hora, formato, unidades, areas, status, observacoes)
                VALUES (?, ?, ?, ?, ?, ?, 'Planejado', '')
            """, (tema, data, hora, formato, unidades_str, areas_str))
            msg = "Treinamento cadastrado com sucesso."

        cronogramas_raw = cur.execute("""
            SELECT c.id, t.nome, c.data, c.hora, c.formato, c.unidades, c.areas, c.status
            FROM cronogramas c
            JOIN temas t ON c.tema_id = t.id
            ORDER BY c.data, c.hora
        """).fetchall()

        cronogramas = []
        for c in cronogramas_raw:
            data_br = datetime.strptime(c[2], "%Y-%m-%d").strftime("%d/%m/%Y")
            cronogramas.append((c[0], c[1], data_br, c[3], c[4], c[5], c[6], c[7]))

        unidades = cur.execute("SELECT * FROM unidades").fetchall()
        areas = cur.execute("SELECT * FROM areas").fetchall()
        temas = cur.execute("SELECT * FROM temas").fetchall()

    return render_template("cronograma.html", msg=msg, cronogramas=cronogramas, unidades=unidades, areas=areas, temas=temas)

# Edição de cronogramas
@app.route("/cronograma/<int:id>", methods=["GET", "POST"])
def editar_cronograma(id):
    with sqlite3.connect(DB) as conn:
        cur = conn.cursor()

        if request.method == "POST":
            status = request.form.get("status")
            observacoes = request.form.get("observacoes")
            cur.execute("""
                UPDATE cronogramas
                SET status = ?, observacoes = ?
                WHERE id = ?
            """, (status, observacoes, id))
            return redirect(url_for("cronograma"))

        c = cur.execute("""
            SELECT c.id, t.nome, c.data, c.hora, c.formato, c.unidades, c.areas, c.status, c.observacoes
            FROM cronogramas c
            JOIN temas t ON c.tema_id = t.id
            WHERE c.id = ?
        """, (id,)).fetchone()

    data_formatada = datetime.strptime(c[2], "%Y-%m-%d").strftime("%d/%m/%Y")
    c = (c[0], c[1], data_formatada, c[3], c[4], c[5], c[6], c[7], c[8])
    return render_template("editar_cronograma.html", c=c)

# ❗Cole aqui:
@app.route("/cronograma/excluir/<int:id>")
def excluir_cronograma(id):
    with sqlite3.connect(DB) as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM cronogramas WHERE id = ?", (id,))
    return redirect(url_for("cronograma"))


# Relatórios
@app.route("/relatorios", methods=["GET"])
def relatorios():
    unidades = request.args.getlist("unidade")
    areas = request.args.getlist("area")
    temas = request.args.getlist("tema")
    data_inicio = request.args.get("data_inicio")
    data_fim = request.args.get("data_fim")
    status = request.args.getlist("status")

    with sqlite3.connect(DB) as conn:
        cur = conn.cursor()

        query = """
            SELECT c.id, t.nome, c.data, c.hora, c.formato, c.unidades, c.areas, c.status
            FROM cronogramas c
            JOIN temas t ON c.tema_id = t.id
            WHERE 1=1
        """
        params = []

        if unidades:
            query += " AND (" + " OR ".join(["c.unidades LIKE ?" for _ in unidades]) + ")"
            params.extend([f"%{u}%" for u in unidades])

        if areas:
            query += " AND (" + " OR ".join(["c.areas LIKE ?" for _ in areas]) + ")"
            params.extend([f"%{a}%" for a in areas])

        if temas:
            query += " AND (" + " OR ".join(["t.nome = ?" for _ in temas]) + ")"
            params.extend(temas)

        if status:
            query += " AND (" + " OR ".join(["c.status = ?" for _ in status]) + ")"
            params.extend(status)

        if data_inicio:
            query += " AND c.data >= ?"
            params.append(data_inicio)

        if data_fim:
            query += " AND c.data <= ?"
            params.append(data_fim)

        cronogramas = cur.execute(query, params).fetchall()
        temas_list = cur.execute("SELECT * FROM temas").fetchall()
        unidades_list = cur.execute("SELECT * FROM unidades").fetchall()
        areas_list = cur.execute("SELECT * FROM areas").fetchall()

    relatorio = []
    for c in cronogramas:
        data_br = datetime.strptime(c[2], "%Y-%m-%d").strftime("%d/%m/%Y")
        relatorio.append((c[0], c[1], data_br, c[3], c[4], c[5], c[6], c[7]))

    return render_template("relatorios.html", relatorio=relatorio, temas=temas_list, unidades=unidades_list, areas=areas_list)


# Exportação para Excel
@app.route("/exportar_xlsx")
def exportar_xlsx():
    with sqlite3.connect(DB) as conn:
        df = pd.read_sql_query("""
            SELECT t.nome as Tema, c.data as Data, c.hora as Hora, c.formato as Formato,
                   c.unidades as Unidades, c.areas as Areas, c.status as Status, c.observacoes as Observacoes
            FROM cronogramas c
            JOIN temas t ON c.tema_id = t.id
        """, conn)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
    output.seek(0)

    return send_file(output, as_attachment=True, download_name="relatorio_treinamentos.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# Exportação para PDF (paisagem + quebra automática)
@app.route("/exportar_pdf")
def exportar_pdf():
    buffer = io.BytesIO()
    styles = getSampleStyleSheet()
    style = ParagraphStyle("TableBody", parent=styles["BodyText"], fontSize=8, leading=10)

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=20,
        leftMargin=20,
        topMargin=30,
        bottomMargin=20
    )

    with sqlite3.connect(DB) as conn:
        data = conn.execute("""
            SELECT t.nome, c.data, c.hora, c.formato, c.unidades, c.areas, c.status, c.observacoes
            FROM cronogramas c
            JOIN temas t ON c.tema_id = t.id
        """).fetchall()

        # Cabeçalho da tabela
        table_data = [[
            Paragraph("<b>Tema</b>", style),
            Paragraph("<b>Data</b>", style),
            Paragraph("<b>Hora</b>", style),
            Paragraph("<b>Formato</b>", style),
            Paragraph("<b>Unidades</b>", style),
            Paragraph("<b>Áreas</b>", style),
            Paragraph("<b>Status</b>", style),
            Paragraph("<b>Observações</b>", style)
        ]]

        # Linhas de dados com quebra automática
        for row in data:
            data_br = datetime.strptime(row[1], "%Y-%m-%d").strftime("%d/%m/%Y")
            row_paragraphs = [Paragraph(str(col), style) for col in
                              (row[0], data_br, row[2], row[3], row[4], row[5], row[6], row[7])]
            table_data.append(row_paragraphs)

        # Definindo larguras personalizadas por coluna (em pontos)
        table = Table(
            table_data,
            repeatRows=1,
            colWidths=[70, 45, 35, 50, 120, 120, 50, 130]  # ajuste proporcional
        )

        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))

    doc.build([Paragraph("Relatório de Treinamentos", styles["Title"]), table])
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="relatorio_treinamentos.pdf",
        mimetype="application/pdf"
    )

@app.route("/excluir/<tipo>/<int:id>")
def excluir(tipo, id):
    with sqlite3.connect(DB) as conn:
        cur = conn.cursor()
        if tipo == "unidade":
            cur.execute("DELETE FROM unidades WHERE id = ?", (id,))
        elif tipo == "area":
            cur.execute("DELETE FROM areas WHERE id = ?", (id,))
        elif tipo == "tema":
            cur.execute("DELETE FROM temas WHERE id = ?", (id,))
    return redirect(url_for("cadastrar"))

@app.route("/editar/<tipo>/<int:id>", methods=["GET", "POST"])
def editar(tipo, id):
    with sqlite3.connect(DB) as conn:
        cur = conn.cursor()

        if request.method == "POST":
            novo_nome = request.form.get("nome")
            if tipo == "unidade":
                nova_sigla = request.form.get("sigla")
                cur.execute("UPDATE unidades SET nome = ?, sigla = ? WHERE id = ?", (novo_nome, nova_sigla, id))
            elif tipo == "area":
                cur.execute("UPDATE areas SET nome = ? WHERE id = ?", (novo_nome, id))
            elif tipo == "tema":
                cur.execute("UPDATE temas SET nome = ? WHERE id = ?", (novo_nome, id))
            return redirect(url_for("cadastrar"))

        if tipo == "unidade":
            dado = cur.execute("SELECT id, nome, sigla FROM unidades WHERE id = ?", (id,)).fetchone()
        elif tipo == "area":
            dado = cur.execute("SELECT id, nome FROM areas WHERE id = ?", (id,)).fetchone()
        elif tipo == "tema":
            dado = cur.execute("SELECT id, nome FROM temas WHERE id = ?", (id,)).fetchone()
        else:
            return redirect(url_for("cadastrar"))

    return render_template("editar.html", tipo=tipo, dado=dado)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=10000)