import streamlit as st
import pandas as pd
import pyautogui
import time

def atualizar_contratos(sap_username, sap_password, ambiente, contrato):
    """
    Atualiza preços cadastrados em contratos na transação ME32K do SAP.
    Replica a ordem de preços pelos "itens" cadastrados na transação. 
    """
    sap_transaction = 'ME32K'
    codigos_contratos = contrato['Contrato'].unique()

    # abrir SAP
    time.sleep(1)
    pyautogui.moveTo(1, 763)
    time.sleep(2)
    pyautogui.click()
    time.sleep(2)
    pyautogui.write('SAP')
    time.sleep(3)
    pyautogui.press('enter')
    time.sleep(5)
        
    # # definir ambiente para abrir
    pyautogui.write(ambiente)
    # time.sleep(1)
    # pyautogui.press('enter')
    # time.sleep(5)
        
    # # logar no SAP
    # pyautogui.write(sap_username)
    # time.sleep(2)
    # pyautogui.press('tab') 
    # time.sleep(2)
    # pyautogui.write(sap_password)
    # time.sleep(2)
    # pyautogui.press('enter')
    # time.sleep(2)
        
    # # buscar transação
    # pyautogui.write(sap_transaction)
    # time.sleep(2)
    # pyautogui.press('enter')
    # time.sleep(2)
        
    # for j in codigos_contratos:
    #     # identificar contratos
    #     pyautogui.write(str(j))
    #     time.sleep(2)
    #     pyautogui.press('enter')
    #     time.sleep(2)
    #     pyautogui.press('enter')
    #     time.sleep(2)

    # # habilitar edição dos preços
    # with pyautogui.hold('ctrl'):
    #         pyautogui.press('f1')
    # time.sleep(1)
    # pyautogui.press('enter')
    # time.sleep(1)
    # with pyautogui.hold('shift'):
    #         pyautogui.press('f6')
    # time.sleep(1)
    # pyautogui.press('enter')
    # time.sleep(1)

    # # atualizar preços
    # for i in contrato[contrato['Contrato'] == j]['Preço líq.']:
    #     pyautogui.press('enter')
    #     time.sleep(2)
    #     pyautogui.press('tab')
    #     time.sleep(1)
    #     pyautogui.hotkey('ctrl', 'a')
    #     pyautogui.press('del')
    #     time.sleep(1)
    #     pyautogui.write(str(i).replace('.', ','))
    #     time.sleep(1)
    #     pyautogui.press('enter')
    #     time.sleep(1)
    #     pyautogui.press('f3')
    #     time.sleep(2)
    #     pyautogui.press('enter')
    #     time.sleep(2)


    # time.sleep(2)
    # with pyautogui.hold('ctrl'):
    #         pyautogui.press('s')
    # time.sleep(2)
    # pyautogui.press('tab')
    # time.sleep(2)
    # pyautogui.press('tab')
    # time.sleep(2)
    # pyautogui.press('enter')
    # time.sleep(2)
    # with pyautogui.hold('ctrl'):
    #         pyautogui.press('a')
    # time.sleep(2)   



def main():

    if 'confirmed' not in st.session_state:
        st.session_state['confirmed'] = False

    st.title('Atualização de preços de contratos SAP')
    st.subheader('Transação ME32K')

    sap_username = st.text_input("SAP Username")
    sap_password = st.text_input("SAP Password", type='password')
    ambiente = st.text_input("Ambiente SAP")
    uploaded_file = st.file_uploader("Inclua o arquivo Excel com os novos preços (de acordo com template)", type=['xlsx'])
    
    if uploaded_file is not None:
        contrato = pd.read_excel(uploaded_file)
        
        if st.button('Atualizar preços'):
            st.session_state['confirmed'] = True
        
        if st.session_state['confirmed']:
            if st.button('VPN está ligada? Se sim, clique aqui'):
                with st.spinner('Atualizando...'):
                    atualizar_contratos(sap_username, sap_password, ambiente, contrato)
                st.success('Preços atualizados!')
                st.session_state['confirmed'] = False 

            if st.button('Cancelar'):
                st.session_state['confirmed'] = False 


if __name__ == "__main__":
    main()
