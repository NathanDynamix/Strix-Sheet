import mongoose from "mongoose";

const connectDB = async () => {
  try {
    const mongoURI = process.env.MONGODB_URI;
    const databaseName = process.env.DATABASE_NAME || "Spreadsheet";
    
    if (!mongoURI) {
      throw new Error("MONGODB_URI environment variable is not defined");
    }

    const connectionInstance = await mongoose.connect(mongoURI);
    console.log(
      `\n MongoDB connected !! DB HOST: ${connectionInstance.connection.host}`
    );
    console.log(`Database Name: ${databaseName}`);
  } catch (error) {
    console.log("MONGODB connection FAILED ", error);
    process.exit(1);
  }
};

export default connectDB;
