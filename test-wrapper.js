const axios = require('axios');

const testTags = ['2PP', '8YG8VQG', 'PYLQRGR'];
const wrapperUrl = 'http://localhost:4000';

async function testWrapperAPI() {
    console.log('Testing Wrapper API...\n');
    
    for (const tag of testTags) {
        try {
            console.log(`Testing tag: ${tag}`);
            const response = await axios.get(`${wrapperUrl}/${tag}`);
            console.log(`✅ Success: ${response.data.name} (TH${response.data.townHallLevel})`);
        } catch (error) {
            console.log(`❌ Failed: ${error.response?.status} - ${error.response?.data?.reason || error.message}`);
        }
        console.log('---');
    }
}

testWrapperAPI();