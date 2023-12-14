const mongoose = require('mongoose');

const studentSchema = new mongoose.Schema({
  name: String,
  mobile: String,
  email: String,
  amount: Number,
  numberoftree: Number,
  certificatePath: String, // Add a field for the certificate file path
});

const Student = mongoose.model('Student', studentSchema);

module.exports = Student;
