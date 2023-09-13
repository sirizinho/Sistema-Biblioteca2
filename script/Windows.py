from PyQt5 import uic,QtWidgets
import Logic as l


app = QtWidgets.QApplication([])

catalogo = uic.loadUi("windows/Catalogo.ui")
minhas_reservas = uic.loadUi("windows/Minhas_Reservas.ui")
reserva = uic.loadUi("windows/Reserva.ui")
visualizar = uic.loadUi("windows/Visualizari.ui")
login = uic.loadUi("windows/Login.ui")
suporte = uic.loadUi("windows/Suporte.ui")
visuTicket = uic.loadUi("windows/VisualizarTicket.ui")
cadAluno = uic.loadUi("windows/Cad_Aluno.ui")
cadLivro = uic.loadUi("windows/Cad_Livro.ui")
login2 = uic.loadUi("windows/Login2.ui")

login.pushButton_2.clicked.connect(login2.show)

login2.pushButton.clicked.connect(l.ValidaLogin2)
login.pushButton.clicked.connect(l.ValidaLogin)

login.show()
app.exec()