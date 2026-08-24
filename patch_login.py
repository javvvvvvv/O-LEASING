# ============================================================================
# PROPIEDAD INTELECTUAL Y LICENCIA COMERCIAL CERRADA
# ============================================================================
# Autor Legal y Titular de Derechos: JAVIER ILLAN GONZALEZ
# Organización: ORANGE CREW
# Contacto: ILLANJAVIER9@GMAIL.COM
#
# ADVERTENCIA LEGAL (MÉXICO Y GLOBAL):
# Este código fuente y su arquitectura son propiedad intelectual exclusiva de
# JAVIER ILLAN GONZALEZ. Queda estrictamente prohibida su reproducción,
# distribución, modificación, ingeniería inversa, copia o uso comercial sin la
# autorización expresa y por escrito del autor. Obra protegida conforme a la
# Ley Federal del Derecho de Autor y tratados internacionales aplicables.
# ============================================================================
import sys

try:
    with open('app.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()

    start_idx = -1
    end_idx = -1
    for i, line in enumerate(lines):
        if 'Bienvenida y login: el video se reproduce desenfocado en el fondo' in line:
            start_idx = i
        if 'st.stop()' in line and start_idx != -1:
            end_idx = i
            break

    if start_idx == -1 or end_idx == -1:
        print('Could not find boundaries')
        sys.exit(1)

    replacement = '''# Bienvenida y login: el video se reproduce desenfocado en el fondo
# y la ventana de login aparece inmediatamente (sin loader).
if not st.session_state.get('auth_user'):
    st.markdown(
        '<style>[data-testid="stSidebar"],[data-testid="stHeader"],footer{display:none!important;} '
        '.main .block-container{padding:0!important;max-width:100%!important;}</style>',
        unsafe_allow_html=True,
    )

    with st.container(key='splash_stage'):
        _video_disponible = __import__('os').path.exists(_VIDEO_BIENVENIDA)
        if _video_disponible:
            try:
                # Video desenfocado de fondo
                st.video(_VIDEO_BIENVENIDA, autoplay=True, muted=True, loop=True)
                st.markdown(
                    """<style>
                    .st-key-splash_stage [data-testid="stVideo"] video {
                        filter: blur(15px) brightness(0.6);
                        transform: scale(1.1); /* Evitar bordes blancos por el blur */
                    }
                    </style>""", unsafe_allow_html=True
                )
            except Exception:
                pass

        # Login Glassmorphism inmediato (animación de pop-in rápido de 0.3s)
        st.markdown(
            f"""<style>
            .st-key-login_wrap {{
                position: absolute;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                z-index: 10;
                background: rgba(255, 255, 255, 0.1);
                backdrop-filter: blur(10px);
                padding: 40px;
                border-radius: 12px;
                border: 1px solid rgba(255, 255, 255, 0.2);
                animation: popInLogin 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.275) forwards;
            }}
            .st-key-login_wrap h3, .st-key-login_wrap p {{
                color: white !important;
            }}
            </style>""",
            unsafe_allow_html=True
        )

        with st.container(key='login_wrap'):
            with st.container():
                _logo_login_path = __import__('os').path.join(BASE_DIR, 'assets', 'o-leasing-logo.png')
                if __import__('os').path.exists(_logo_login_path):
                    st.image(_logo_login_path, width=190)
                if not hay_usuarios(AUTH_DB):
                    st.markdown('<h3>Crea el usuario administrador</h3>', unsafe_allow_html=True)
                    st.markdown('<p>Es la primera vez que se abre el sistema.</p>', unsafe_allow_html=True)
                    with st.form('bootstrap_admin'):
                        _bu = st.text_input('Usuario')
                        _bn = st.text_input('Nombre completo')
                        _bp1 = st.text_input('Contraseña', type='password')
                        _bp2 = st.text_input('Confirmar contraseña', type='password')
                        if st.form_submit_button('Crear administrador', use_container_width=True):
                            if _bp1 != _bp2:
                                st.error('Las contraseñas no coinciden.')
                            else:
                                ok, msg = crear_usuario(AUTH_DB, _bu, _bn, _bp1, 'admin')
                                if ok:
                                    st.success(msg + ' Ahora inicia sesión.')
                                    st.rerun()
                                else:
                                    st.error(msg)
                else:
                    st.markdown('<h3>Iniciar sesión</h3>', unsafe_allow_html=True)
                    st.markdown('<p>Sistema de gestión de arrendamiento</p>', unsafe_allow_html=True)
                    with st.form('login_form'):
                        _lu = st.text_input('Usuario')
                        _lp = st.text_input('Contraseña', type='password')
                        if st.form_submit_button('Entrar', use_container_width=True):
                            _user = verificar_login(AUTH_DB, _lu, _lp)
                            if _user:
                                st.session_state['auth_user'] = _user
                                st.rerun()
                            else:
                                st.error('Usuario o contraseña incorrectos.')
    st.stop()
\n'''

    new_lines = lines[:start_idx] + [replacement] + lines[end_idx+1:]
    with open('app.py', 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    print('Login updated successfully')
except Exception as e:
    print('Error:', e)
