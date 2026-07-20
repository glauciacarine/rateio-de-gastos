import streamlit as st
from collections import defaultdict
import pandas as pd

st.title("💸 Rateio de Despesas em Grupo")

# Participantes
st.header("👥 Participantes")
if "pessoas" not in st.session_state:
    st.session_state.pessoas = []

apelido = st.text_input("Apelido do participante")
if st.button("Adicionar participante"):
    if apelido and apelido not in st.session_state.pessoas:
        st.session_state.pessoas.append(apelido)

st.write("Participantes adicionados:", st.session_state.pessoas)

# Gastos
st.header("🧾 Gastos")
if "gastos" not in st.session_state:
    st.session_state.gastos = []

col1, col2, col3 = st.columns(3)
with col1:
    pagador = st.selectbox("Quem pagou?", st.session_state.pessoas)
with col2:
    categoria = st.text_input("Categoria do gasto")
with col3:
    valor = st.number_input("Valor do gasto", min_value=0.0, format="%.2f")

participantes_gasto = st.multiselect("Quem participou desse gasto?", st.session_state.pessoas)

if st.button("Adicionar gasto"):
    if pagador and valor > 0 and categoria and participantes_gasto:
        st.session_state.gastos.append((pagador, valor, categoria, participantes_gasto))

# Mostrar lista de gastos com opção de exclusão
if st.session_state.gastos:
    st.subheader("📋 Lista de gastos")
    for i, (pagador, valor, categoria, participantes_gasto) in enumerate(st.session_state.gastos):
        col1, col2 = st.columns([4, 1])
        with col1:
            participantes_str = ", ".join(participantes_gasto)
            st.write(f"{pagador} → {categoria} → R${valor:.2f} (dividido entre: {participantes_str})")
        with col2:
            if st.button("❌", key=f"del_{i}"):
                st.session_state.gastos.pop(i)
                st.experimental_rerun()

# Cálculo do rateio
if st.button("Calcular rateio"):
    pessoas = st.session_state.pessoas
    gastos = st.session_state.gastos

    if not pessoas or not gastos:
        st.error("Adicione participantes e gastos antes de calcular.")
    else:
        pagamentos = defaultdict(float)
        consumos = defaultdict(float)
        total_gasto = 0

        for pagador, valor, categoria, participantes_gasto in gastos:
            pagamentos[pagador] += valor
            total_gasto += valor
            valor_por_pessoa = valor / len(participantes_gasto)
            for pessoa in participantes_gasto:
                consumos[pessoa] += valor_por_pessoa

        saldos = {p: pagamentos[p] - consumos[p] for p in pessoas}

        st.subheader("📊 Resultado do Rateio")
        st.write(f"Total gasto: R${total_gasto:.2f}")

        st.write("### Saldos individuais:")
        for p in pessoas:
            saldo = saldos[p]
            status = "deve" if saldo < 0 else "tem a receber"
            st.write(f"{p}: R${abs(saldo):.2f} ({status})")

        st.write("### Sugestões de pagamento:")
        devedores = {p: -s for p, s in saldos.items() if s < 0}
        credores = {p: s for p, s in saldos.items() if s > 0}

        pagamentos_sugeridos = []
        for devedor, valor_devido in devedores.items():
            for credor in list(credores.keys()):
                if valor_devido == 0:
                    break
                valor_a_receber = credores[credor]
                pagamento = min(valor_devido, valor_a_receber)
                st.write(f"{devedor} deve pagar R${pagamento:.2f} para {credor}")
                pagamentos_sugeridos.append({
                    "Devedor": devedor,
                    "Credor": credor,
                    "Valor": round(pagamento, 2)
                })
                valor_devido -= pagamento
                credores[credor] -= pagamento

        # Botão para salvar em Excel
        if st.button("💾 Salvar dados em Excel"):
            df_participantes = pd.DataFrame({"Participantes": pessoas})
            df_gastos = pd.DataFrame([
                {
                    "Pagador": g[0],
                    "Valor": g[1],
                    "Categoria": g[2],
                    "Participantes": ", ".join(g[3])
                } for g in gastos
            ])
            df_rateio = pd.DataFrame([
                {
                    "Pessoa": p,
                    "Pagou": pagamentos[p],
                    "Consumiu": consumos[p],
                    "Saldo": saldos[p],
                    "Status": "deve" if saldos[p] < 0 else "tem a receber"
                } for p in pessoas
            ])
            df_pagamentos = pd.DataFrame(pagamentos_sugeridos)

            with pd.ExcelWriter("rateio_despesas.xlsx", engine="openpyxl") as writer:
                df_participantes.to_excel(writer, sheet_name="Participantes", index=False)
                df_gastos.to_excel(writer, sheet_name="Gastos", index=False)
                df_rateio.to_excel(writer, sheet_name="Rateio", index=False)
                if not df_pagamentos.empty:
                    df_pagamentos.to_excel(writer, sheet_name="Pagamentos sugeridos", index=False)

            st.success("Arquivo 'rateio_despesas.xlsx' salvo com sucesso!")

# Botão para limpar todos os dados
if st.button("🔄 Limpar tudo"):
    st.session_state.pessoas = []
    st.session_state.gastos = []
    st.experimental_rerun()