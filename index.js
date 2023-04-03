const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const cors = require("cors");
const Excel = require('exceljs');
const { MongoClient } = require("mongodb");



const app = express();
const port = process.env.PORT || 3000;

const url = "mongodb+srv://dasddj2:MN0K791ZmzdL6Qyi@cluster2.a0fbnc2.mongodb.net/?retryWrites=true&w=majority";
const client = new MongoClient(url);
// The database to use
const dbName = "bdozio";
const db = client.db(dbName);
//var actualizar=false;


app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());



//PERMITIR CORS
  
// enabling CORS for any unknow origin(https://xyz.example.com)
app.use(cors());
//CONECTAR A BD
async function run() {
  try {
      await client.connect();
      console.log("Conectado a la base de datos");
  } catch (err) {
      console.log(err.stack);
  }
  finally {
      await client.close();
  }
}
//


app.get('/productos-consulta', async (req, res) => {
  try {
    await client.connect();
    console.log("Connected correctly to server");
   

    // Use the collection "departamentos"
    const col = db.collection("productos");

    // Find all documents in the collection
    const cursor = col.find();
    // Convert cursor to array
    const result = await cursor.toArray();  
    // Send response with array of documents
    res.send(result);

  } catch (err) {
    console.log(err.stack);
    res.status(500).send('Internal server error');
  } finally {
    await client.close();
  }
});

app.get('/descargar', async (req, res) => {
  try {
    await client.connect();
    console.log("CONEXION! DESCARGAR EXCEL");

    // Use the collection "productos"
    const col = db.collection("productos");

    // Find all documents in the collection
    const cursor = col.find();
    // Convert cursor to array
    const result = await cursor.toArray();

    // Crear el archivo Excel
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Productos');
    worksheet.columns = [
      { header: 'Código', key: 'Codigo', width: 15 },
      { header: 'Descripción', key: 'Descripcion', width: 30 },
      { header: 'Precio Costo', key: 'PrecioCosto', width: 15 },
      { header: 'Precio Venta', key: 'PrecioVenta', width: 15 },
      { header: 'Precio Mayoreo', key: 'PrecioMayoreo', width: 15 },
      { header: 'Existencias', key: 'Existencias', width: 20 },
      { header: 'Departamento', key: 'Departamento', width: 20 }
    ];
    worksheet.addRows(result);

    // Enviar el archivo Excel como descarga al cliente
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Productos.xlsx');
    await workbook.xlsx.write(res);

  } catch (err) {
    console.log(err.stack);
    res.status(500).send('Internal server error');
  } finally {
    await client.close();
  }
});

app.post('/insertar-productos', async (req, res) => {

 const insertarProductos = req.body;

  try {
    await client.connect();
    const collection = client.db('bdozio').collection('productos');
    const result = await collection.insertOne(insertarProductos);
    res.status(201).send('Insertado');
    console.log('Insertado');

  } catch (err) {
    console.log(err);
    res.status(500).send('Error interno del servidor');
  } finally {
    await client.close();
  }
  
});

app.post('/consulta/', async (req, res) => { 

  const cod = parseInt(req.body.codigo_actualizar);

  try {
    await client.connect();
    const collection = client.db('bdozio').collection('productos');

    // Verificamos si el producto existe en la base de datos
    const consulta = await collection.findOne({ Codigo: cod });
    if (!consulta) {
      res.status(404).send('El producto no existe');
      console.log('El producto no existe');
      return;
    }
    else{
      actualizar=1;
      console.log('Existe');
    }
    

  } catch (err) {
    console.log(err);
    res.status(500).send('Error interno del servidor');
  } finally {
    await client.close();
  }


 
});

app.put('/actualizar-productos/:id',  async (req, res) => {

  const cod = parseInt(req.params.id);
  const updateFields = req.body;

  try {
 
    await client.connect();
       // Actualizar el documento con los campos específicos
    
    db.collection("productos").updateOne({Codigo: cod}, {$set: updateFields})
    .then(result => {
    res.send("Documento actualizado");
  })
  .catch(err => {
    console.log("Error actualizando documento: ", err);
    res.status(500).send("Error actualizando documento");
  });
  } catch (err) {
    console.log(err);
    res.status(500).json({ message: 'Error interno del servidor' });
  } finally {
    //await client.close();
    setTimeout(() => {client.close()}, 1500)
  }

});




app.delete('/eliminar-productos/:codigo', async (req, res) => {
  const codigo = parseInt(req.params.codigo);

  try {
    await client.connect();
    const collection = client.db('bdozio').collection('productos');

    // Verificamos si el producto existe en la base de datos
    const consulta = await collection.findOne({ Codigo: codigo });
    if (!consulta) {
      res.status(404).send('El producto no existe');
      console.log('El producto no existe');
      return;
    }

    // Eliminamos el producto de la base de datos
    const result = await collection.deleteOne({ Codigo: codigo });
    res.status(200).send('Eliminado');
    console.log('Eliminado');

  } catch (err) {
    console.log(err);
    res.status(500).send('Error interno del servidor');
  } finally {
    await client.close();
  }
});

app.get('/productos-consulta-parametro/:codigo', async (req, res) => {
  const codigo = parseInt(req.params.codigo);

  try {
    await client.connect();
    const collection = client.db('bdozio').collection('productos');

    // Buscamos el producto por su código
    const producto = await collection.findOne({ Codigo: codigo });

    if (!producto) {
      res.status(404).send('Producto no encontrado');
      return;
    }

    res.send(producto);

  } catch (err) {
    console.log(err);
    res.status(500).send('Error interno del servidor');
  } finally {
    await client.close();
  }
});





app.listen(port, function() {
  console.log(`Servidor iniciado en http://localhost:${port}`);
});
run();