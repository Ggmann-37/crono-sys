# crono-sys

## Protección de autenticación con Cloudflare Turnstile

El formulario de inicio de sesión y registro usa el widget visible **Managed** de Cloudflare Turnstile. Solo aparece en esas dos vistas de autenticación; la página principal y la vista previa no lo muestran. La site key pública está en `js/app.js`; el token generado por el widget se envía a Supabase mediante `captchaToken` y Supabase valida el token usando la clave secreta configurada en **Authentication → Attack Protection → CAPTCHA protection**.

La clave secreta nunca debe añadirse al repositorio ni al código del navegador. En Supabase, selecciona **Turnstile by Cloudflare**, pega la secret key y guarda los cambios. Si el desafío no se completa o el token caduca, el formulario no permite continuar y Supabase rechaza igualmente cualquier token inválido.