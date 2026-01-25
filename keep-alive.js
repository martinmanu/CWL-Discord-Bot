const cron = require('node-cron');
const express = require('express');

// Create a simple health check endpoint
const app = express();
app.get('/health', (req, res) => {
    res.json({ status: 'alive', timestamp: new Date().toISOString() });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Health check server running on port ${PORT}`);
});

// Self-ping every 14 minutes to keep Render service awake
cron.schedule('*/14 * * * *', async () => {
    try {
        const response = await fetch(`https://your-render-app.onrender.com/health`);
        console.log(`Keep-alive ping: ${response.status} at ${new Date().toISOString()}`);
    } catch (error) {
        console.log(`Keep-alive failed: ${error.message} at ${new Date().toISOString()}`);
    }
});

console.log('Keep-alive cron job started - self-pinging every 14 minutes');