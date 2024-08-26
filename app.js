const express = require('express');
const {
    initialize,
    getUsers,
    addUser,
    updateUser,
    deleteUser
} = require('./services/excelManipulation.service.js');

const app = express();
const PORT = 3000;

app.use(express.json());

app.get('/users', (req, res) => {
    res.json(getUsers());
});

app.post('/users', (req, res) => {
    const { firstName, lastName, age } = req.body;

    if (!firstName || !lastName || !age) {
        return res.status(400).json({ message: 'Please provide all required fields: firstName, lastName, age.' });
    }

    addUser(firstName, lastName, age)
        .then(newUser => {
            res.status(201).json(newUser);
        })
        .catch(err => {
            res.status(500).json({ message: err.message });
        });
});

app.put('/users/:firstName', (req, res) => {
    const { firstName } = req.params;
    const { lastName, age } = req.body;

    updateUser(firstName, { lastName, age })
        .then(updatedUser => {
            res.json(updatedUser);
        })
        .catch(err => {
            res.status(500).json({ message: err.message });
        });
});

app.delete('/users/:firstName', (req, res) => {
    const { firstName } = req.params;

    deleteUser(firstName)
        .then(() => {
            res.status(204).send();
        })
        .catch(err => {
            res.status(500).json({ message: err.message });
        });
});

initialize().then(() => {
    app.listen(PORT, () => {
        console.log(`Server is running on http://localhost:${PORT}`);
    });
});
package.json