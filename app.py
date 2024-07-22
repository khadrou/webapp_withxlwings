from flask import Flask, jsonify, request
import xlwings as xw
import numpy as np
import os 
os.environ["XLWINGS_LICENSE_KEY"]="eyJwcm9kdWN0cyI6IFsicHJvIiwgInJlcG9ydHMiXSwgInZhbGlkX3VudGlsIjogIjIwMjQtMDgtMTciLCAibGljZW5zZV90eXBlIjogInRyaWFsIn0=9e0ab"

app = Flask(__name__)

class Pipe:
    def __init__(self, diameter, length, fluid_density, fluid_viscosity, singularities):
        self.diameter = diameter
        self.length = length
        self.fluid_density = fluid_density
        self.fluid_viscosity = fluid_viscosity
        self.singularities = singularities

    def calculate_pressure_drop(self):
        singularities = {
            'Elbow 90°': 0.16,
            'Elbow 45°': 0.09,
            'Elbow 30°': 0.4,
            'T-Junction': 1.2,
            'Y-Junction': 1.5,
            'Diffuser': 2.0,
            'Orifice': 3.0
        }

        def friction_factor(Re, ε, D):
            return (1 / (2 * np.log(Re))) + (ε / (3.7 * D))

        Re = (self.diameter * self.fluid_density * self.length) / self.fluid_viscosity
        ε = 0.0001  # roughness coefficient
        D = self.diameter
        f = friction_factor(Re, ε, D)

        pressure_drop = 0
        for sing in self.singularities:
            if sing in singularities:
                K = singularities[sing]
                pressure_drop += K * (f * self.length / D)
            else:
                print(f"Invalid singularity: {sing}")

        return pressure_drop

class PipeNetwork:
    def __init__(self, pipes):
        self.pipes = pipes

    def calculate_total_pressure_drop(self):
        total_pressure_drop = 0
        for pipe in self.pipes:
            total_pressure_drop += pipe.calculate_pressure_drop()
        return total_pressure_drop

@app.get('/')
def index():
    return 'Serveur calcul 2024 merciii'

@app.post('/calculate_total_pressure_drop')
def calculate_total_pressure_drop():
    try:
        print('running')
        print(request.json)
        with xw.Book(json=request.json) as book:
            sht_input = book.sheets[0]
            print('===>', sht_input['A1'].value)

            pipe1_diameter_cell = sht_input.range('B1')
            pipe1_length_cell = sht_input.range('B2')
            pipe1_fluid_density_cell = sht_input.range('B3')
            pipe1_fluid_viscosity_cell = sht_input.range('B4')
            pipe1_singularities_cell = sht_input.range('B5')
            pipe1_singularities = [sing.strip() for sing in pipe1_singularities_cell.value.split(',')]

            pipe2_diameter_cell = sht_input.range('C1')
            pipe2_length_cell = sht_input.range('C2')
            pipe2_fluid_density_cell = sht_input.range('C3')
            pipe2_fluid_viscosity_cell = sht_input.range('C4')
            pipe2_singularities_cell = sht_input.range('C5')
            pipe2_singularities = [sing.strip() for sing in pipe2_singularities_cell.value.split(',')]

            pipe3_diameter_cell = sht_input.range('D1')
            pipe3_length_cell = sht_input.range('D2')
            pipe3_fluid_density_cell = sht_input.range('D3')
            pipe3_fluid_viscosity_cell = sht_input.range('D4')
            pipe3_singularities_cell = sht_input.range('D5')
            pipe3_singularities = [sing.strip() for sing in pipe3_singularities_cell.value.split(',')]

            pipe1 = Pipe(pipe1_diameter_cell.value, pipe1_length_cell.value, pipe1_fluid_density_cell.value, pipe1_fluid_viscosity_cell.value, pipe1_singularities)
            pipe2 = Pipe(pipe2_diameter_cell.value, pipe2_length_cell.value, pipe2_fluid_density_cell.value, pipe2_fluid_viscosity_cell.value, pipe2_singularities)
            pipe3 = Pipe(pipe3_diameter_cell.value, pipe3_length_cell.value, pipe3_fluid_density_cell.value, pipe3_fluid_viscosity_cell.value, pipe3_singularities)

            pipe_network = PipeNetwork([pipe1, pipe2, pipe3])

            total_pressure_drop = pipe_network.calculate_total_pressure_drop()
            print(total_pressure_drop)
            sht_input["E7"].value = total_pressure_drop
            sht_input["A7"].value = "Done !!"
 
            return book.json()
    except Exception as e:
        print('Error: ', e)
        return jsonify({'actions': 0})
        print(total_pressure_drop)
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5500, debug=True)
