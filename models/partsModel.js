const mongoose = require('mongoose');

const Schema=mongoose.Schema;

const partsSchema = new Schema({
   
model:{
    type: String
},
customer:{
type:String
},
defectType:{
    type:String
},
defectSubType:{
    type:String
},
ok:{
    type:Number
},
rework:{
    type:Number
},
reject:{
    type:Number
}
},{timestamps:true})

module.exports=mongoose.model("Part", partsSchema);