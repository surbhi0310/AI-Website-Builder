import mongoose from "mongoose";
import dotenv from "dotenv";
import dns from "dns";

dotenv.config();

// Use public DNS servers for MongoDB SRV records
dns.setServers(["8.8.8.8", "1.1.1.1"]);

let memoryServerInstance = null;

const Databaseconnect = async () => {
  if (process.env.MONGO_URL) {
    try {
      await mongoose.connect(process.env.MONGO_URL, {
        serverSelectionTimeoutMS: 3000,
      });
      console.log("✅ Database connected successfully to MongoDB Atlas/URI");
      return;
    } catch (error) {
      console.warn(
        `⚠️ MongoDB Atlas connection failed (${error.message}). Starting local in-memory MongoDB...`
      );
    }
  }

  try {
    const { MongoMemoryServer } = await import("mongodb-memory-server");
    memoryServerInstance = await MongoMemoryServer.create();
    const uri = memoryServerInstance.getUri();
    await mongoose.connect(uri);
    console.log(`✅ Connected successfully to local in-memory MongoDB at: ${uri}`);
  } catch (err) {
    console.error("❌ Failed to connect to MongoDB:", err.message);
  }
};

export default Databaseconnect;
