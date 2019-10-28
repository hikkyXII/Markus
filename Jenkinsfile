#!groovy
properties([disableConcurrentBuilds()])

pipeline {
    agent {
        label 'windows && python'
    }
    options {
        buildDiscarder(logRotator(numToKeepStr: '10', artifactNumToKeepStr: '10'))
        timestamps()
    }
    stages {
        stage("Install requirements") {
            steps {
                bat 'pip install -r requirements.txt'
            }
        }
        stage("Build") {
            steps {
                bat 'pyinstaller -F markus.py'
            }
        }
    }

    post {
        always {
            archiveArtifacts artifacts: 'dist\\markus.exe', fingerprint: true
        }
    }
}