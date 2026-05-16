pipeline {
    agent any

    parameters {
        choice(
            name: 'SCRIPT_NUMBER',
            choices: ['all', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'],
            description: 'Seleccione el número de script a ejecutar. Usa "all" para ejecutar todos los scripts.'
        )
    }

    environment {
        VENV_DIR = '/var/jenkins_home/workspace/Generacion_Actas/venv'
        // Establece la política CSP vacía para permitir que Jenkins muestre correctamente el HTML incrustado 
        JAVA_OPTS = "-Dhudson.model.DirectoryBrowserSupport.CSP=\"sandbox allow-scripts allow-same-origin; default-src 'none'; img-src 'self' data:; style-src 'self' 'unsafe-inline' data:; script-src 'self' 'unsafe-inline' 'unsafe-eval';\""
        // Establece las variables de allure
        APP_VERSION = '1.0.0'
        PLATFORM = 'Fedora Linux 41 (Server Edition)'
        BROWSER = 'Chromedriver: 128.0.6613.84'
        ALLURE_RESULTS_DIR = 'reports/allure-results'
    }
    stages {
        stage('Clean Up and Checkout ') {
            steps {
                deleteDir()
                //Clonar el repositorio Git
                git url: 'https://github.com/ericruizINE/Generacion_Actas.git', branch: 'main'
            }
        }
        stage('Install & Setup venv') {
            steps {
                sh "python3 -m venv ${VENV_DIR}"
            }
        }
        stage('Install Dependencies') {
            steps {
                // Activar el entorno virtual e instalar las dependencias
                sh """
                    . ${VENV_DIR}/bin/activate > /dev/null 2>&1
                    pip install --no-cache-dir -r requirements.txt
                    pip install --no-cache-dir allure-pytest
                """
            }
        }
        stage('Preparar ambiente') {
            steps {
                script {
                    // Generar archivo environment.properties con variables de entorno
                    def alluredir = env.ALLURE_RESULTS_DIR
                    sh "mkdir -p ${alluredir}"
                    sh """
                        echo 'APP_VERSION=${env.APP_VERSION}' >> ${alluredir}/environment.properties
                        echo 'PLATFORM=${env.PLATFORM}' >> ${alluredir}/environment.properties
                        echo 'BROWSER=${env.BROWSER}' >> ${alluredir}/environment.properties
                        echo 'BUILD_URL=${env.BUILD_URL}' >> ${alluredir}/environment.properties
                    """
                }
            }
        }
        stage('Run generation scripts') {
            steps {
                script {
                    def scriptMap = [
                        '1': '01_actaspresidencia.py',
                        '2': '02_actaspresidencia_especial.py',
                        '3': '03_actassenadurias.py',
                        '4': '04_actassenaduriasrp.py',
                        '5': '05_actasdiputaciones.py',
                        '6': '06_actasdiputacionesrp.py',
                        '7': '07_actaspresidenciava.py',
                        '8': '08_actaspresidenciave.py',
                        '9': '09_actaspresidenciavpp.py',
                        '10': '10_actasenaduriasva.py',
                        '11': '11_actassenaduriave.py',
                        '12': '12_actasdiputacionesva.py'
                    ]

                    def selectedScripts = []
                    if (params.SCRIPT_NUMBER == 'all') {
                        selectedScripts = scriptMap.values().toList()
                    } else if (scriptMap.containsKey(params.SCRIPT_NUMBER)) {
                        selectedScripts = [scriptMap[params.SCRIPT_NUMBER]]
                    } else {
                        error "Valor de SCRIPT_NUMBER inválido: ${params.SCRIPT_NUMBER}"
                    }

                    if (isUnix()) {
                        for (scriptFile in selectedScripts) {
                            sh "${VENV_DIR}/bin/python ${scriptFile}"
                        }
                    } else {
                        for (scriptFile in selectedScripts) {
                            bat "python ${scriptFile}"
                        }
                    }
                }
            }
        }

        stage('Archive artifacts') {
            steps {
                archiveArtifacts artifacts: 'Actas/**', fingerprint: true
            }
        }
    }

    post {
        always {
            script {
                // Ejecuta Allure
                allure includeProperties: false, jdk: '', reportBuildPolicy: 'ALWAYS', results: [[path: "${env.ALLURE_RESULTS_DIR}"]]
                
                // Define las URLs de los reportes
                def allureReportUrl = "${env.BUILD_URL}allure"

                env.BUILD_RESULT = currentBuild.currentResult
                // Convertir la duración a un formato legible
                def durationMillis = currentBuild.duration
                def durationSeconds = (durationMillis / 1000) as int
                def minutes = (durationSeconds / 60) as int
                def seconds = durationSeconds % 60
                env.BUILD_DURATION = "${minutes}m ${seconds}s"

                // Imprime las URLs en consola
                echo "El reporte de Allure está disponible en: ${allureReportUrl}"
                
            }
        }
    }
}