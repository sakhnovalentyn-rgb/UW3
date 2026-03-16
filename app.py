import streamlit as st
import io
from main import process_document  # Імпорт вашої великої функції
import time

st.set_page_config(page_title="Кредитний Аналіз", layout="centered")
st.title("📂 Автоматизація кредитних висновків")

# --- ФУНКЦІЯ ПЕРЕВІРКИ ПАРОЛЯ ---
def check_password():
    """Повертає True, якщо ПІБ введено та пароль вірний."""
    st.sidebar.title("Авторизація")
    
    # Додаємо поле для ПІБ
    user_fullname = st.sidebar.text_input("Введіть ПІБ аналітика", placeholder="Іванов І.І.")
    
    # Поле для пароля
    password = st.sidebar.text_input("Введіть пароль", type="password")
    
    if st.sidebar.button("Увійти"):
        if not user_fullname.strip():
            st.sidebar.error("Будь ласка, введіть ПІБ!")
            return False, None
        
        if password == "UW2026": # Замініть на ваш реальний пароль
            st.session_state["authenticated"] = True
            st.session_state["user_name"] = user_fullname
            return True, user_fullname
        else:
            st.sidebar.error("Невірний пароль")
            return False, None
    
    # Перевірка, чи вже залогінені
    if st.session_state.get("authenticated", False):
        return True, st.session_state.get("user_name")
    
    return False, None

# --- ГОЛОВНИЙ ІНТЕРФЕЙС ---
is_auth, analyst_name = check_password()

if is_auth:
    st.title("💳 Система автоматичного андеррайтингу")
    st.write(f"Вітаємо, **{analyst_name}**! Ви можете розпочати роботу.")
    
    # Завантаження файлу
    uploaded_file = st.file_uploader("Завантажте ваш .docx файл", type=["docx"])

    if uploaded_file is not None:
        
        # Показуємо кнопку лише після завантаження файлу
        if st.button("🚀 Обробити та підготувати звіт"):
            start_time = time.time()
            
            # Зберігаємо ПІБ, але не очищуємо всю сесію, щоб не вилетіти з логіну
            # Замість clear() видаляємо лише попередні результати, якщо вони були
            if 'processed_data' in st.session_state:
                del st.session_state['processed_data']

            with st.spinner('Проводжу розрахунки фінансових показників...'):
                # ВАЖЛИВО: Передаємо analyst_name у функцію обробки
                result_file, data = process_document(uploaded_file, analyst_name)
                
                duration = time.time() - start_time
                
                if result_file:
                    st.success("✅ Аналіз завершено успішно!")
                    st.info(f"⏱ Час обробки файлу: {duration:.2f} сек.")
                    
                    # Кнопка скачування
                    st.download_button(
                        label="📥 Скачати заповнений .docx",
                        data=result_file,
                        file_name=f"Висновок_андерайтера_{analyst_name}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error("❌ Сталася помилка під час обробки. Перевірте консоль VS Code.")
else:
    st.info("Будь ласка, введіть дані в бічній панелі (sidebar) для доступу до системи.")