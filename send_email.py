import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def enviar_correo():
    # Configuración del servidor SMTP de Gmail
    smtp_host = "correo.ine.mx"
    remitente = "pruebasQA@ine.mx"
    smtp_user = os.getenv('SMTP_USER')
    smtp_password = os.getenv('SMTP_PASS')

    build_name = os.getenv('JOB_NAME', 'Desconocido')
    build_result = os.getenv('BUILD_RESULT', 'Desconocido')
    build_duration = os.getenv('BUILD_DURATION', 'Desconocido')
    build_number = os.getenv('BUILD_NUMBER', 'Desconocido')
    build_url = os.getenv('BUILD_URL', 'Desconocido')
    blue_ocean_url = f"{os.getenv('JENKINS_URL', '')}blue/organizations/jenkins/{build_name}/detail/{build_name}/{build_number}/pipeline"

    destinatarios_env = os.getenv('DESTINATARIOS', 'eric.ruiz@ine.mx')
    destinatarios = [email.strip() for email in destinatarios_env.split(',') if email.strip()]

    script_executed = os.getenv('SCRIPT_EXECUTED', 'No disponible')
    artifact_urls_raw = os.getenv('ARTIFACT_URLS', 'No disponible')
    artifact_urls = [url.strip() for url in artifact_urls_raw.split(',') if url.strip()]

    subject = f"[DEST][CI/CD] Resultado de ejecución de Pipeline: {build_name} Número: {build_number}"

    artifact_links_html = '<li>No se encontraron artefactos.</li>'
    if artifact_urls and artifact_urls != ['No disponible']:
        artifact_links_html = ''.join([
            f"<li><a href='{url}'>{url.rstrip('/').split('artifact/')[-1]}</a></li>"
            for url in artifact_urls
        ])

    body = f"""
        <h2 style="color: #2E86C1;">Reporte de Ejecución del Pipeline</h2>
        <p>Estimado equipo,</p>
        <p>El pipeline <strong>{build_name}</strong> ha finalizado.</p>
        <table style="width: 100%; border: 1px solid #ddd; border-collapse: collapse;">
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd;">Build Number</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{build_number}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd;">Estado</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{build_result}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd;">Duración</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{build_duration}</td>
            </tr>
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd;">Script ejecutado</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{script_executed}</td>
            </tr>
        </table>
        <p>Enlaces de interés:</p>
        <ul>
            <li><a href="{build_url}">Pipeline Jenkins</a></li>
            <li><a href="{blue_ocean_url}">Pipeline Blue Ocean</a></li>
        </ul>
        <p><strong>Artefactos disponibles:</strong></p>
        <ul>
            {artifact_links_html}
        </ul>
        <p>Atentamente,<br>Equipo DevOps</p>
    """

    mensaje = MIMEMultipart()
    mensaje['From'] = remitente
    mensaje['To'] = ', '.join(destinatarios)
    mensaje['Subject'] = subject
    mensaje.attach(MIMEText(body, 'html'))

    try:
        with smtplib.SMTP(smtp_host, 587) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(smtp_user, smtp_password)
            server.sendmail(remitente, destinatarios, mensaje.as_string())
            print('✅ Correo enviado con éxito.')
            print(f'Destinatarios: {destinatarios}')
            print(f'Artifact URLs: {artifact_urls}')
    except Exception as e:
        print(f'Error al enviar el correo: {e}')


if __name__ == '__main__':
    enviar_correo()
