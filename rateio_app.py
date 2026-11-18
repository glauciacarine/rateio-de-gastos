import io
from collections import defaultdict

import pandas as pd
import streamlit as st

st.title("💸 Rateio de Despesas em Grupo")

# Margem para comparar valores em dinheiro (evita problemas de arredondamento do float)
TOLERANCIA = 0.005

# ---------------------------------------------------------------
# Estado da sessão (memória entre um clique e outro)
# ---------------------------------------------------------------
if "pessoas" not in st.session_state:
    st.session_state.pessoas = []
if "gastos" not in st.session_state:
    st.session_state.gastos = []
if "resultado" not in st.session_state:
    st.session_state.resultado = None  # guarda o último cálculo do rateio


def limpar_resultado():
    """Sempre que os dados mudam, o cálculo anterior deixa de valer."""
    st.session_state.resultado = None


# ---------------------------------------------------------------
# Participantes
# ---------------------------------------------------------------
st.header("👥 Participantes")

apelido = st.text_input("Apelido do participante").strip()
if st.button("Adicionar participante"):
    if not apelido:
        st.warning("Informe um apelido.")
    elif apelido in st.session_state.pessoas:
        st.warning("Esse participante já foi adicionado.")
    else:
        st.session_state.pessoas.append(apelido)
        limpar_resultado()

st.write("Participantes adicionados:", st.session_state.pessoas)

# ---------------------------------------------------------------
# Gastos
# ---------------------------------------------------------------
st.header("🧾 Gastos")

col1, col2, col3 = st.columns(3)
with col1:
    pagador = st.selectbox("Quem pagou?", st.session_state.pessoas)
with col2:
    categoria = st.text_input("Categoria do gasto").strip()
with col3:
    valor = st.number_input("Valor do gasto", min_value=0.0, format="%.2f")

participantes_gasto = st.multiselect(
    "Quem participou desse gasto?", st.session_state.pessoas
)

if st.button("Adicionar gasto"):
    if pagador and valor > 0 and categoria and participantes_gasto:
        st.session_state.gastos.append(
            (pagador, round(valor, 2), categoria, participantes_gasto)
        )
        limpar_resultado()
    else:
        st.warning("Preencha quem pagou, categoria, valor e participantes.")

# Lista de gastos com opção de exclusão
if st.session_state.gastos:
    st.subheader("📋 Lista de gastos")
    for i, (pag, val, cat, parts) in enumerate(st.session_state.gastos):
        c1, c2 = st.columns([4, 1])
        with c1:
            st.write(f"{pag} → {cat} → R${val:.2f} (dividido entre: {', '.join(parts)})")
        with c2:
            if st.button("❌", key=f"del_{i}"):
                st.session_state.gastos.pop(i)
                limpar_resultado()
                st.rerun()


# ---------------------------------------------------------------
# Cálculo do rateio
# ---------------------------------------------------------------
def calcular_rateio(pessoas, gastos):
    pagamentos = defaultdict(float)
    consumos = defaultdict(float)
    total_gasto = 0.0

    for pag, val, _cat, parts in gastos:
        pagamentos[pag] += val
        total_gasto += val
        valor_por_pessoa = val / len(parts)
        for pessoa in parts:
            consumos[pessoa] += valor_por_pessoa

    saldos = {p: round(pagamentos[p] - consumos[p], 2) for p in pessoas}

    # Sugestões de pagamento
    devedores = {p: -s for p, s in saldos.items() if s < -TOLERANCIA}
    credores = {p: s for p, s in saldos.items() if s > TOLERANCIA}

    sugestoes = []
    for devedor, valor_devido in devedores.items():
        for credor in credores:
            if valor_devido <= TOLERANCIA:
                break
            if credores[credor] <= TOLERANCIA:
                continue
            pagamento = min(valor_devido, credores[credor])
            sugestoes.append(
                {"Devedor": devedor, "Credor": credor, "Valor": round(pagamento, 2)}
            )
            valor_devido -= pagamento
            credores[credor] -= pagamento

    return {
        "total": round(total_gasto, 2),
        "pagamentos": dict(pagamentos),
        "consumos": dict(consumos),
        "saldos": saldos,
        "sugestoes": sugestoes,
    }


def status_do_saldo(saldo):
    if saldo > TOLERANCIA:
        return "tem a receber"
    if saldo < -TOLERANCIA:
        return "deve"
    return "quitado"


if st.button("Calcular rateio"):
    if not st.session_state.pessoas or not st.session_state.gastos:
        st.error("Adicione participantes e gastos antes de calcular.")
    else:
        st.session_state.resultado = calcular_rateio(
            st.session_state.pessoas, st.session_state.gastos
        )

# O resultado fica salvo na sessão, então continua na tela
# mesmo depois de outros cliques (como o de baixar o Excel).
resultado = st.session_state.resultado
if resultado:
    pessoas = st.session_state.pessoas
    gastos = st.session_state.gastos

    st.subheader("📊 Resultado do Rateio")
    st.write(f"Total gasto: R${resultado['total']:.2f}")

    st.write("### Saldos individuais:")
    for p in pessoas:
        saldo = resultado["saldos"][p]
        st.write(f"{p}: R${abs(saldo):.2f} ({status_do_saldo(saldo)})")

    st.write("### Sugestões de pagamento:")
    if resultado["sugestoes"]:
        for s in resultado["sugestoes"]:
            st.write(f"{s['Devedor']} deve pagar R${s['Valor']:.2f} para {s['Credor']}")
    else:
        st.write("Ninguém precisa pagar nada. Todos estão quites!")

    # Monta o Excel em memória e oferece para download
    df_participantes = pd.DataFrame({"Participantes": pessoas})
    df_gastos = pd.DataFrame(
        [
            {"Pagador": g[0], "Valor": g[1], "Categoria": g[2], "Participantes": ", ".join(g[3])}
            for g in gastos
        ]
    )
    df_rateio = pd.DataFrame(
        [
            {
                "Pessoa": p,
                "Pagou": round(resultado["pagamentos"].get(p, 0.0), 2),
                "Consumiu": round(resultado["consumos"].get(p, 0.0), 2),
                "Saldo": resultado["saldos"][p],
                "Status": status_do_saldo(resultado["saldos"][p]),
            }
            for p in pessoas
        ]
    )
    df_sugestoes = pd.DataFrame(resultado["sugestoes"])

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_participantes.to_excel(writer, sheet_name="Participantes", index=False)
        df_gastos.to_excel(writer, sheet_name="Gastos", index=False)
        df_rateio.to_excel(writer, sheet_name="Rateio", index=False)
        if not df_sugestoes.empty:
            df_sugestoes.to_excel(writer, sheet_name="Pagamentos sugeridos", index=False)

    st.download_button(
        label="💾 Baixar dados em Excel",
        data=buffer.getvalue(),
        file_name="rateio_despesas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------------------------------------------------------
# Limpar tudo
# ---------------------------------------------------------------
if st.button("🔄 Limpar tudo"):
    st.session_state.pessoas = []
    st.session_state.gastos = []
    limpar_resultado()
    st.rerun()