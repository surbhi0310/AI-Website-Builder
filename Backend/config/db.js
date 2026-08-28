import mongoose from "mongoose";
import dotenv from "dotenv";
import dns from "dns";

dotenv.config();

// Use public DNS servers because the current network DNS
// doesn't resolve MongoDB SRV records correctly for Node.js.
dns.setServers(["8.8.8.8", "1.1.1.1"]);

const Databaseconnect = async () => {
  try {
    await mongoose.connect(process.env.MONGO_URL);
    console.log("Database connected successfully");
  } catch (error) {
    console.log(
      `This is error from the mongodb connection: ${error.message}`
    );
  }
};

export default Databaseconnect;