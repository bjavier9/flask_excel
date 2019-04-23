from flask import Flask, jsonify, request
from flask_sqlalchemy import SQLAlchemy
import flask_excel as excel

app = Flask(__name__)

app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///tmp.db'
db = SQLAlchemy(app)
class Post(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cod_materia = db.Column(db.String(80))
    nom_materia = db.Column(db.String(80))
    cod_grupo = db.Column(db.Integer)
    ID_TIPO_GRUPO = db.column(db.Integer)
    TIPO_GRUPO = db.column(db.String(80))
    cod_est_grupo = db.column(db.Integer)
    estado = db.column(db.String(80))
    cupo = db.column(db.Integer)
    disponibles = db.column(db.Integer)
    matriculados = db.column(db.Integer)
    prematriculados = db.column(db.Integer)
    levantamientos = db.column(db.Integer)
    disponibles_real = db.column(db.Integer)
    horariogrp = db.column(db.String(80))
    nom_periodo = db.column(db.String(80))
    cod_edificio = db.column(db.String(80))
    cod_escuela = db.column(db.Integer)
    escuela = db.column(db.String(80))
    entra = db.column(db.String(80))
    sale = db.column(db.String(80))
    dia = db.column(db.String(80))
    cod_aula = db.column(db.Integer)
    TIPO_DOCENTE = db.column(db.String(80))
    NOMBRE = db.column(db.String(80))
    PROFESOR = db.column(db.String(80))

    def __init__(self, cod_materia,  nom_materia, cod_grupo,ID_TIPO_GRUPO, 
    TIPO_GRUPO, cod_est_grupo, estado, cupo, disponibles, matriculados, 
    prematriculados,levantamientos, disponibles_real, horariogrp,
    nom_periodo, cod_edificio, escuela, entra, sale, dia, cod_aula,
    TIPO_DOCENTE, NOMBRE, PROFESOR  
    ):
        self.cod_materia = cod_materia
        self.nom_materia =  nom_materia
        self.cod_grupo = cod_grupo
        self.ID_TIPO_GRUPO = ID_TIPO_GRUPO
        self.TIPO_GRUPO = TIPO_GRUPO
        self.cod_est_grupo = cod_est_grupo
        self.estado = estado
        self.cupo = cupo
        self.disponibles = disponibles
        self.matriculados = matriculados
        self.levantamientos = levantamientos
        self.disponibles_real = disponibles_real
        self.horariogrp = horariogrp
        self.nom_periodo = nom_periodo
        self.cod_edificio = cod_edificio
        self.cod_escuela = cod_escuela
        self.escuela = escuela
        self.entra = entra
        self.sale = sale
        self.dia = dia
        self.cod_aula = cod_aula
        self.TIPO_DOCENTE = TIPO_DOCENTE
        self.NOMBRE = NOMBRE
        self.PROFESOR = PROFESOR

    def __repr__(self):
        return '<Post %r>' % self.cod_materia


@app.route("/import", methods=['GET', 'POST'])
def importoexcel():
    if request.method == 'POST':

        def post_init_func(row):           
            p = Post(row['cod_materia'],row['nom_materia'], row['cod_grupo'],row['ID_TIPO_GRUPO'],row['TIPO_GRUPO']
            ,row['cod_est_grupo'],row['estado'],row['cupo'], row['disponibles'], row['matriculados'], row['levantamientos']
            ,row['disponibles_real'], row['horariogrp'], row['nom_periodo'], row['cod_edificio'], row['cod_escuela']
            ,row['entra'], row['sale'], row['dia'], row['cod_aula'], row['TIPO_DOCENTE'], row['NOMBRE'], row['PROFESOR'])
            return p
        request.save_book_to_database(
            field_name='file', session=db.session,
            tables=[Post],
            initializers=[post_init_func])
        return redirect(url_for('.handson_table'), code=302)
    return '''
    <!doctype html>
    <title>Upload an excel file</title>
    <h1>Excel file upload (xls, xlsx, ods please)</h1>
    <form action="" method=post enctype=multipart/form-data><p>
    <input type=file name=file><input type=submit value=Upload>
    </form>
    '''
@app.route("/export", methods=['GET'])
def exportoexcel():
    return excel.make_response_from_tables(db.session, [Post], "xlsx")

@app.route("/u1/cursos/", methods=['GET'])
def cursos():
    cursos =  Post.query.order_by(Post.id).all()
    return jsonify({
        "data": [{"id": x.id, "materia": x.nom_materia,} for x in cursos]}),200

@app.route("/curso/<item>", methods=["GET"])
def busqueda_curso(item):    
    post =  Post.query.filter(Post.cod_grupo==item).all()
    return jsonify({
      "data": [{"id": x.id, "materia": x.nom_materia, "cod_materia": x.cod_grupo, "disponibles":x.disponibles,"matriculados":x.matriculados,"cod_edificio":x.cod_edificio,"Profesor":x.PROFESOR} for x in post]
    })       

if __name__ == "__main__":
    excel.init_excel(app)
    db.create_all()
    app.run(debug=True)

    