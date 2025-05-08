const express = require('express')
const app = express()
const port = 3000

app.use(express.json())


app.post('/msp/issue', (req, res) => {
    const inputString = req.body

    // if (!inputString) {
	// console.log('Status 400: Нет строки для обработки');
    //     return res.status(400).json({ error: 'Нет строки для обработки' });
    // }
    //
    console.log(inputString);
    //
    // // Возвращаем массив в формате JSON
     return res.status(201).json(inputString);
    //return inputString
})


app.listen(3000)
//     port, () => {
//   console.log(`Example app listening on port ${port}`)
// })

