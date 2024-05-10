var mssql=require("mssql");
const dotenv = require('dotenv');
const path = require("path");
const { env } = process;
require('dotenv').config({
  path: path.resolve(
      __dirname,
      `./env.${process.env.NODE_ENV ? process.env.NODE_ENV : "test"}`
    ),
});
const config = {
  user: env.USER,
  password: env.PASSWORD,
  server: env.HOST, // You can use 'localhost\\instance' to connect to named instance
  database: env.DATABASE,
  options: {
      encrypt: false, // Use this if you're on Windows Azure
      enableArithAbort: true
  }
};
  let pool = new mssql.ConnectionPool(config);
  var query=function(sql,options,callback){
    pool.connect().then(()=>{
        pool.request().query(sql,options,function(err,results,fields){
            if(callback!=undefined) callback(err,results,fields);
            pool.close();
        })
    });
    
    
  }
  module.exports=query;