pipeline{
    agent{
        label 'any'
    }

    options{
        timeout( time: 5, unit: 'MINUTES' )
    }

    stages{

        stage('Checkout'){
            steps{
                // git branch: 'master', url: 'https://alkjsdfh.com'
                // sh 'echo hello'
                echo 'hello'
            }
        }
        stage('Build'){
            steps{
                // sh 'echo Build'
                echo 'Build'
            }
        }
        stage('Test'){
            steps{
                // sh 'echo Test'
                echo 'Test'
            }
        }
        stage('Approve'){
            steps{
                script{
                    def userInput= input message: 'Proceed for deployment',
                    ok:'Yes',
                    parameters:[]
                    echo "Approved"
                }    
            }
        }
        stage('Deploy'){
            steps{
                // sh 'echo Deploy'
                echo 'Deploy'
            }
        }
        stage('Rollback'){
            steps{
                sh 'echo Rollback'
            }
        }
    }
    post{
        success{
            echo 'ok'
        }
        failure{
            echo 'fail'
        }
        always{
            echo 'clean'
        }
    }
}