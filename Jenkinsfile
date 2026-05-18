pipeline {
    agent any

    parameters {
        choice(
            name: 'SCRIPT_NUMBER',
            choices: [
                'all',
                '1 - Presidencia',
                '2 - Presidencia especial',
                '3 - Senadurias',
                '4 - Senadurias RP',
                '5 - Diputaciones',
                '6 - Diputaciones RP',
                '7 - Presidencia VA',
                '8 - Presidencia VE',
                '9 - Presidencia VPP',
                '10 - Senadurias VA',
                '11 - Senadurias VE',
                '12 - Diputaciones VA'
            ],
            description: 'Seleccione el script a ejecutar. Use "all" para ejecutar todos los scripts.'
        )
    }

    environment {
        VENV_DIR = '/var/jenkins_home/workspace/Generacion_Actas/venv'
        // Establece la política CSP vacía para permitir que Jenkins muestre correctamente el HTML incrustado 
        JAVA_OPTS = "-Dhudson.model.DirectoryBrowserSupport.CSP=\"sandbox allow-scripts allow-same-origin; default-src 'none'; img-src 'self' data:; style-src 'self' 'unsafe-inline' data:; script-src 'self' 'unsafe-inline' 'unsafe-eval';\""
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
                """
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
                    def selectionKey = params.SCRIPT_NUMBER == 'all' ? 'all' : params.SCRIPT_NUMBER.split(' - ')[0]
                    if (selectionKey == 'all') {
                        selectedScripts = scriptMap.values().toList()
                    } else if (scriptMap.containsKey(selectionKey)) {
                        selectedScripts = [scriptMap[selectionKey]]
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
                archiveArtifacts artifacts: 'Archivos/Datos/**', fingerprint: true
            }
        }
    }

    post {
        always {
            script {
                env.BUILD_RESULT = currentBuild.currentResult
                def durationMillis = currentBuild.duration
                def durationSeconds = (durationMillis / 1000) as int
                def minutes = (durationSeconds / 60) as int
                def seconds = durationSeconds % 60
                env.BUILD_DURATION = "${minutes}m ${seconds}s"
                echo "Resultado del build: ${env.BUILD_RESULT}"
                echo "Duración del build: ${env.BUILD_DURATION}"
            }
        }
    }
}