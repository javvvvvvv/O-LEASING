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

replacement = '''# Bienvenida y login: el video se reproduce desenfocado en el fondo
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
                st.video(_VIDEO_BIENVENIDA, autoplay=True, muted=True, loop=True)
                st.markdown(
                    """<style>
                    /* Video as fullscreen background */
                    .st-key-splash_stage [data-testid="stVideo"] {
                        position: fixed !important;
                        top: 50% !important;
                        left: 50% !important;
                        min-width: 100% !important;
                        min-height: 100% !important;
                        width: auto !important;
                        height: auto !important;
                        transform: translate(-50%, -50%) scale(1.1) !important;
                        z-index: 0 !important;
                    }
                    .st-key-splash_stage [data-testid="stVideo"] video {
                        object-fit: cover !important;
                        filter: blur(8px) brightness(0.4) !important;
                    }
                    </style>""", unsafe_allow_html=True
                )
            except Exception:
                pass

        st.markdown("<div style='height: 12vh;'></div>", unsafe_allow_html=True)
        col_L, col_C, col_R = st.columns([1, 1.2, 1])
        
        with col_C:
            st.markdown(
                """<style>
                /* Opaque box for login */
                .st-key-login_wrap {
                    position: relative;
                    z-index: 10;
                    background-color: var(--background-color, #ffffff) !important;
                    padding: 40px;
                    border-radius: 12px;
                    box-shadow: 0 10px 40px rgba(0, 0, 0, 0.4);
                    border: 1px solid rgba(150,150,150,0.3);
                }
                .st-key-login_wrap h3, .st-key-login_wrap p {
                    text-align: center;
                }
                </style>""", unsafe_allow_html=True
            )
            with st.container(key='login_wrap'):
                _logo_login_path = __import__('os').path.join(BASE_DIR, 'assets', 'o-leasing-logo.png')
                if __import__('os').path.exists(_logo_login_path):
                    col_logo1, col_logo2, col_logo3 = st.columns([1, 2, 1])
                    with col_logo2:
                        st.image(_logo_login_path, use_container_width=True)
                if not hay_usuarios(AUTH_DB):
                    st.markdown('<h3 style="text-align:center;">Crea el usuario administrador</h3>', unsafe_allow_html=True)
                    st.markdown('<p style="text-align:center;">Es la primera vez que se abre el sistema.</p>', unsafe_allow_html=True)
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
                    st.markdown('<h3 style="text-align:center;">Iniciar sesión</h3>', unsafe_allow_html=True)
                    st.markdown('<p style="text-align:center;">Sistema de gestión de arrendamiento</p>', unsafe_allow_html=True)
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
'''

with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

login_start = -1
login_end = -1
for i, line in enumerate(lines):
    if 'Bienvenida y login' in line and 'el video se reproduce desenfocado' in line:
        login_start = i
    if 'st.stop()' in line and login_start != -1:
        login_end = i
        break

if login_start != -1:
    lines = lines[:login_start] + [replacement + '\n'] + lines[login_end+1:]
    with open('app.py', 'w', encoding='utf-8') as f:
        f.writelines(lines)
    print('Login fixed')
else:
    print('Not found')
