require('dotenv').config();
const { Client, GatewayIntentBits, SlashCommandBuilder, EmbedBuilder, AttachmentBuilder } = require('discord.js');
const axios = require('axios');
const ExcelJS = require('exceljs');

const client = new Client({ intents: [GatewayIntentBits.Guilds] });
const sessions = new Map();
const playerCache = new Map();
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

class Player {
    constructor(tag, name, townHall, heroes, equipment, pets) {
        this.tag = tag;
        this.name = name;
        this.townHall = townHall;
        this.heroes = heroes;
        this.equipment = equipment;
        this.pets = pets;
        this.score = this.calculateScore();
    }

    calculateScore() {
        const heroTotal = this.heroes.king + this.heroes.queen + this.heroes.warden + this.heroes.champion + this.heroes.minion;
        const equipmentTotal = this.equipment.reduce((sum, level) => sum + level, 0);
        const petTotal = this.pets.reduce((sum, level) => sum + level, 0);
        return (this.townHall * 1000) + (heroTotal * 10) + (equipmentTotal * 5) + (petTotal * 5);
    }
}

async function fetchPlayerData(tag) {
    const cleanTag = tag.replace('#', '').toUpperCase();
    const cacheKey = `player_${cleanTag}`;
    
    if (playerCache.has(cacheKey)) {
        const cached = playerCache.get(cacheKey);
        if (Date.now() - cached.timestamp < 300000) {
            return cached.data;
        }
    }

    try {
        const response = await axios.get(`https://api.clashofclans.com/v1/players/%23${cleanTag}`, {
            headers: { 'Authorization': `Bearer ${process.env.CLASH_API_KEY}` }
        });

        const player = response.data;
        const heroes = {
            king: player.heroes?.find(h => h.name === 'Barbarian King')?.level || 0,
            queen: player.heroes?.find(h => h.name === 'Archer Queen')?.level || 0,
            warden: player.heroes?.find(h => h.name === 'Grand Warden')?.level || 0,
            champion: player.heroes?.find(h => h.name === 'Royal Champion')?.level || 0,
            minion: player.heroes?.find(h => h.name === 'Minion Prince')?.level || 0
        };

        const equipment = [];
        if (player.heroEquipment) {
            player.heroEquipment.forEach(eq => {
                if (eq.village === 'home') equipment.push(eq.level || 0);
            });
        }

        const pets = [];
        if (player.troops) {
            const petNames = ['L.A.S.S.I', 'Mighty Yak', 'Electro Owl', 'Unicorn', 'Phoenix', 'Poison Lizard', 'Diggy', 'Frosty', 'Spirit Fox', 'Angry Jelly', 'Sneezy', 'Meteor Golem'];
            player.troops.forEach(troop => {
                if (troop.village === 'home' && petNames.includes(troop.name)) {
                    pets.push(troop.level || 0);
                }
            });
        }

        const playerData = new Player(`#${cleanTag}`, player.name, player.townHallLevel, heroes, equipment, pets);
        
        playerCache.set(cacheKey, {
            data: playerData,
            timestamp: Date.now()
        });

        return playerData;
    } catch (error) {
        if (error.response?.status === 429) {
            await delay(1000);
            return fetchPlayerData(tag);
        }
        throw new Error(`Failed to fetch player ${tag}: ${error.message}`);
    }
}

async function generateExcel(players, assignments = {}, filename = 'cwl_roster') {
    const workbook = new ExcelJS.Workbook();
    const clanGroups = {};
    
    players.forEach(player => {
        const clan = assignments[player.tag] || 'Unassigned';
        if (!clanGroups[clan]) clanGroups[clan] = [];
        clanGroups[clan].push(player);
    });

    const colors = [
        { header: 'FF2E4B8A', row: 'FFE8F0FF' },
        { header: 'FF8B4513', row: 'FFFFF8DC' },
        { header: 'FF228B22', row: 'FFF0FFF0' },
        { header: 'FF800080', row: 'FFFAF0E6' },
        { header: 'FFDC143C', row: 'FFFFEFD5' },
        { header: 'FF4682B4', row: 'FFF0F8FF' },
        { header: 'FF8B008B', row: 'FFFFE4E1' },
        { header: 'FF2F4F4F', row: 'FFF5F5F5' }
    ];

    let colorIndex = 0;
    Object.entries(clanGroups).forEach(([clanName, clanPlayers]) => {
        const worksheet = workbook.addWorksheet(clanName);
        const colorScheme = colors[colorIndex % colors.length];
        
        worksheet.columns = [
            { header: 'Player', key: 'name', width: 20 },
            { header: 'Tag', key: 'tag', width: 15 },
            { header: 'TH', key: 'th', width: 5 },
            { header: 'Hero Total', key: 'heroTotal', width: 12 },
            { header: 'Equipment Total', key: 'equipmentTotal', width: 15 },
            { header: 'Pet Total', key: 'petTotal', width: 12 },
            { header: 'Score', key: 'score', width: 10 }
        ];

        worksheet.getRow(1).eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colorScheme.header } };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });

        let clanTotalScore = 0;
        clanPlayers.forEach((player) => {
            const heroTotal = player.heroes.king + player.heroes.queen + player.heroes.warden + player.heroes.champion + player.heroes.minion;
            const equipmentTotal = player.equipment.reduce((sum, level) => sum + level, 0);
            const petTotal = player.pets.reduce((sum, level) => sum + level, 0);
            clanTotalScore += player.score;
            
            const row = worksheet.addRow({
                name: player.name,
                tag: player.tag,
                th: player.townHall,
                heroTotal: heroTotal,
                equipmentTotal: equipmentTotal,
                petTotal: petTotal,
                score: player.score
            });

            row.eachCell((cell) => {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colorScheme.row } };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            });
        });

        const totalRow = worksheet.addRow({
            name: 'TOTAL CLAN WEIGHT',
            tag: '', th: '', heroTotal: '', equipmentTotal: '', petTotal: '',
            score: clanTotalScore
        });
        
        totalRow.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colorScheme.header } };
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });

        colorIndex++;
    });

    const buffer = await workbook.xlsx.writeBuffer();
    return { buffer, filename };
}

const commands = [
    new SlashCommandBuilder()
        .setName('start')
        .setDescription('Load player tags for CWL management')
        .addStringOption(option =>
            option.setName('tags')
                .setDescription('Player tags separated by spaces, commas, or newlines')
                .setRequired(true)
        ),
    new SlashCommandBuilder()
        .setName('rank')
        .setDescription('Show players ranked by strength (strongest first)'),
    new SlashCommandBuilder()
        .setName('sortequally')
        .setDescription('Split players into two balanced teams using snake draft'),
    new SlashCommandBuilder()
        .setName('cwl')
        .setDescription('Distribute players across CWL clans')
        .addIntegerOption(option =>
            option.setName('size')
                .setDescription('Players per clan (15 or 30)')
                .setRequired(true)
                .addChoices(
                    { name: '15v15', value: 15 },
                    { name: '30v30', value: 30 }
                )
        )
        .addIntegerOption(option =>
            option.setName('clans')
                .setDescription('Number of clans')
                .setRequired(true)
        )
        .addStringOption(option =>
            option.setName('names')
                .setDescription('Clan names separated by commas (optional)')
                .setRequired(false)
        ),
    new SlashCommandBuilder()
        .setName('sort')
        .setDescription('Show players sorted by strength (weakest first)'),
    new SlashCommandBuilder()
        .setName('export')
        .setDescription('Export roster data to Excel file')
        .addStringOption(option =>
            option.setName('filename')
                .setDescription('Custom filename for the Excel file (optional)')
                .setRequired(false)
        ),
    new SlashCommandBuilder()
        .setName('clearcache')
        .setDescription('Clear all cached player data and sessions'),
    new SlashCommandBuilder()
        .setName('distribute')
        .setDescription('Distribute players with custom war sizes and sorting')
        .addStringOption(option =>
            option.setName('sizes')
                .setDescription('War sizes separated by commas (e.g., 15,30,15)')
                .setRequired(true)
        )
        .addStringOption(option =>
            option.setName('sort')
                .setDescription('Sorting method')
                .setRequired(true)
                .addChoices(
                    { name: 'Strongest First (desc)', value: 'desc' },
                    { name: 'Weakest First (asc)', value: 'asc' },
                    { name: 'Balanced (equal)', value: 'equal' }
                )
        )
        .addStringOption(option =>
            option.setName('names')
                .setDescription('Clan names separated by commas (optional)')
                .setRequired(false)
        )
];

client.once('ready', async () => {
    console.log(`Logged in as ${client.user.tag}!`);
    try {
        console.log('Registering slash commands...');
        await client.application.commands.set(commands);
        console.log('Slash commands registered successfully!');
        console.log('Commands:', commands.map(c => c.name));
    } catch (error) {
        console.error('Error registering commands:', error);
    }
});

client.on('interactionCreate', async interaction => {
    if (!interaction.isChatInputCommand()) return;

    const { commandName, guildId } = interaction;

    try {
        switch (commandName) {
            case 'start':
                const tags = interaction.options.getString('tags');
                const tagList = tags.split(/[\s,\n]+/).filter(tag => tag.trim());
                
                await interaction.deferReply();
                
                const players = [];
                const errors = [];

                for (let i = 0; i < tagList.length; i++) {
                    try {
                        const player = await fetchPlayerData(tagList[i]);
                        players.push(player);
                        await delay(200);
                    } catch (error) {
                        errors.push(`${tagList[i]}: ${error.message}`);
                    }
                }

                sessions.set(guildId, players);

                const embed = new EmbedBuilder()
                    .setTitle('Players Loaded')
                    .setDescription(`Successfully loaded ${players.length} players`)
                    .setColor(0x00AE86);

                if (errors.length > 0) {
                    embed.addFields({ name: 'Errors', value: errors.slice(0, 5).join('\n') });
                }

                await interaction.editReply({ embeds: [embed] });
                break;

            case 'rank':
                const rankPlayers = sessions.get(guildId);
                if (!rankPlayers) {
                    await interaction.reply('No players loaded. Use /start first.');
                    return;
                }

                const sorted = [...rankPlayers].sort((a, b) => b.score - a.score);
                
                let rankTable = '```\n';
                rankTable += 'Rank| Name        | Tag      | Score | Heroes | Equipment | Pets\n';
                rankTable += '----|-------------|----------|-------|--------|-----------|------\n';
                
                sorted.forEach((p, i) => {
                    const heroTotal = p.heroes.king + p.heroes.queen + p.heroes.warden + p.heroes.champion + p.heroes.minion;
                    const equipTotal = p.equipment.reduce((sum, level) => sum + level, 0);
                    const petTotal = p.pets.reduce((sum, level) => sum + level, 0);
                    
                    rankTable += `${(i + 1).toString().padStart(4)} | ${p.name.substring(0, 17).padEnd(17)} | ${p.tag.padEnd(10)} | ${p.score.toString().padStart(5)} | ${heroTotal.toString().padStart(6)} | ${equipTotal.toString().padStart(9)} | ${petTotal.toString().padStart(4)}\n`;
                });
                
                rankTable += '```';

                const rankEmbed = new EmbedBuilder()
                    .setTitle(`Player Rankings (All ${sorted.length} Players)`)
                    .setDescription(rankTable)
                    .setColor(0x0099FF);

                await interaction.reply({ embeds: [rankEmbed] });
                break;

            case 'sortequally':
                const equalPlayers = sessions.get(guildId);
                if (!equalPlayers) {
                    await interaction.reply('No players loaded. Use /start first.');
                    return;
                }

                const sortedEqual = [...equalPlayers].sort((a, b) => b.score - a.score);
                const teamA = [];
                const teamB = [];

                sortedEqual.forEach((player, index) => {
                    const cycle = Math.floor(index / 2) % 2;
                    if (index % 2 === 0) {
                        if (cycle === 0) teamA.push(player);
                        else teamB.push(player);
                    } else {
                        if (cycle === 0) teamB.push(player);
                        else teamA.push(player);
                    }
                });

                const teamAScore = teamA.reduce((sum, p) => sum + p.score, 0);
                const teamBScore = teamB.reduce((sum, p) => sum + p.score, 0);

                const equalEmbed = new EmbedBuilder()
                    .setTitle('Equal Teams')
                    .addFields(
                        { 
                            name: `Team A (${teamAScore})`, 
                            value: teamA.map(p => `${p.name} (${p.score})`).join('\n') || 'Empty',
                            inline: true 
                        },
                        { 
                            name: `Team B (${teamBScore})`, 
                            value: teamB.map(p => `${p.name} (${p.score})`).join('\n') || 'Empty',
                            inline: true 
                        }
                    )
                    .setColor(0x00FF00);

                await interaction.reply({ embeds: [equalEmbed] });
                break;

            case 'cwl':
                const size = interaction.options.getInteger('size');
                const clans = interaction.options.getInteger('clans');
                const clanNames = interaction.options.getString('names');
                const cwlPlayers = sessions.get(guildId);
                
                if (!cwlPlayers) {
                    await interaction.reply('No players loaded. Use /start first.');
                    return;
                }

                const sortedCwl = [...cwlPlayers].sort((a, b) => b.score - a.score);
                const assignments = {};
                const clanRosters = {};
                
                const cwlCustomNames = clanNames ? clanNames.split(',').map(name => name.trim()) : [];
                for (let i = 0; i < clans; i++) {
                    const clanName = cwlCustomNames[i] || `Clan ${i + 1}`;
                    clanRosters[clanName] = [];
                }

                sortedCwl.forEach((player, index) => {
                    const clanIndex = Math.floor(index / size);
                    if (clanIndex < clans) {
                        const clanName = cwlCustomNames[clanIndex] || `Clan ${clanIndex + 1}`;
                        clanRosters[clanName].push(player);
                        assignments[player.tag] = clanName;
                    } else {
                        assignments[player.tag] = 'Bench';
                    }
                });

                const cwlEmbed = new EmbedBuilder()
                    .setTitle(`CWL ${size}v${size} Distribution`)
                    .setColor(0xFF6600);

                Object.entries(clanRosters).forEach(([clanName, roster]) => {
                    if (roster.length > 0) {
                        const totalScore = roster.reduce((sum, p) => sum + p.score, 0);
                        cwlEmbed.addFields({
                            name: `${clanName} (${totalScore})`,
                            value: roster.slice(0, 10).map(p => `${p.name} (${p.score})`).join('\n'),
                            inline: true
                        });
                    }
                });

                const cwlBenchCount = sortedCwl.length - (size * clans);
                if (cwlBenchCount > 0) {
                    cwlEmbed.addFields({
                        name: `Bench (${cwlBenchCount} players)`,
                        value: `${cwlBenchCount} players on bench`,
                        inline: false
                    });
                }

                sessions.set(`${guildId}_assignments`, assignments);
                await interaction.reply({ embeds: [cwlEmbed] });
                break;

            case 'sort':
                const sortPlayers = sessions.get(guildId);
                if (!sortPlayers) {
                    await interaction.reply('No players loaded. Use /start first.');
                    return;
                }

                const ascending = [...sortPlayers].sort((a, b) => a.score - b.score);
                const sortList = ascending.slice(0, 20).map((p, i) => 
                    `${i + 1}. ${p.name} (TH${p.townHall}) - ${p.score}`
                ).join('\n');

                const sortEmbed = new EmbedBuilder()
                    .setTitle('Players (Weakest First)')
                    .setDescription(sortList)
                    .setColor(0xFF0000);

                await interaction.reply({ embeds: [sortEmbed] });
                break;

            case 'distribute':
                const sizes = interaction.options.getString('sizes').split(',').map(s => parseInt(s.trim()));
                const sortMethod = interaction.options.getString('sort');
                const names = interaction.options.getString('names');
                const distPlayers = sessions.get(guildId);
                
                if (!distPlayers) {
                    await interaction.reply('No players loaded. Use /start first.');
                    return;
                }

                let sortedDist;
                if (sortMethod === 'asc') {
                    sortedDist = [...distPlayers].sort((a, b) => a.score - b.score);
                } else if (sortMethod === 'desc') {
                    sortedDist = [...distPlayers].sort((a, b) => b.score - a.score);
                } else {
                    sortedDist = [...distPlayers].sort((a, b) => b.score - a.score);
                }

                const distAssignments = {};
                const distClanRosters = {};
                const customNames = names ? names.split(',').map(name => name.trim()) : [];
                
                sizes.forEach((size, i) => {
                    const clanName = customNames[i] || `Clan ${i + 1} (${size}v${size})`;
                    distClanRosters[clanName] = [];
                });

                if (sortMethod === 'equal') {
                    sortedDist.forEach((player, index) => {
                        const clanIndex = index % sizes.length;
                        const clanName = customNames[clanIndex] || `Clan ${clanIndex + 1} (${sizes[clanIndex]}v${sizes[clanIndex]})`;
                        if (distClanRosters[clanName].length < sizes[clanIndex]) {
                            distClanRosters[clanName].push(player);
                            distAssignments[player.tag] = clanName;
                        } else {
                            distAssignments[player.tag] = 'Bench';
                        }
                    });
                } else {
                    let playerIndex = 0;
                    sizes.forEach((size, clanIndex) => {
                        const clanName = customNames[clanIndex] || `Clan ${clanIndex + 1} (${size}v${size})`;
                        for (let i = 0; i < size && playerIndex < sortedDist.length; i++) {
                            distClanRosters[clanName].push(sortedDist[playerIndex]);
                            distAssignments[sortedDist[playerIndex].tag] = clanName;
                            playerIndex++;
                        }
                    });
                    
                    while (playerIndex < sortedDist.length) {
                        distAssignments[sortedDist[playerIndex].tag] = 'Bench';
                        playerIndex++;
                    }
                }

                const distEmbed = new EmbedBuilder()
                    .setTitle(`War Distribution (${sortMethod.toUpperCase()} sort)`)
                    .setColor(0xFF6600);

                Object.entries(distClanRosters).forEach(([clanName, roster]) => {
                    if (roster.length > 0) {
                        const totalScore = roster.reduce((sum, p) => sum + p.score, 0);
                        distEmbed.addFields({
                            name: `${clanName} (${totalScore})`,
                            value: roster.slice(0, 10).map(p => `${p.name} (${p.score})`).join('\n'),
                            inline: true
                        });
                    }
                });

                const totalAssigned = sizes.reduce((sum, size) => sum + size, 0);
                const distBenchCount = sortedDist.length - totalAssigned;
                if (distBenchCount > 0) {
                    distEmbed.addFields({
                        name: `Bench (${distBenchCount} players)`,
                        value: `${distBenchCount} players on bench`,
                        inline: false
                    });
                }

                sessions.set(`${guildId}_assignments`, distAssignments);
                await interaction.reply({ embeds: [distEmbed] });
                break;

            case 'clearcache':
                sessions.clear();
                playerCache.clear();
                
                const clearEmbed = new EmbedBuilder()
                    .setTitle('Cache Cleared')
                    .setDescription('All player data and session data has been cleared.')
                    .setColor(0x00FF00);
                
                await interaction.reply({ embeds: [clearEmbed] });
                break;

            case 'export':
                const exportPlayers = sessions.get(guildId);
                if (!exportPlayers) {
                    await interaction.reply('No players loaded. Use /start first.');
                    return;
                }

                await interaction.deferReply();

                const assignments_key = sessions.get(`${guildId}_assignments`) || {};
                const sortedExport = [...exportPlayers].sort((a, b) => b.score - a.score);
                const customFilename = interaction.options.getString('filename') || 'cwl_roster';
                
                const { buffer, filename } = await generateExcel(sortedExport, assignments_key, customFilename);
                const attachment = new AttachmentBuilder(buffer, { name: `${filename}.xlsx` });

                await interaction.editReply({ 
                    content: 'Here is your CWL roster export:',
                    files: [attachment] 
                });
                break;
        }
    } catch (error) {
        console.error('Command error:', error);
        const errorMessage = interaction.deferred ? 
            { content: `Error: ${error.message}` } : 
            `Error: ${error.message}`;
        
        if (interaction.deferred) {
            await interaction.editReply(errorMessage);
        } else {
            await interaction.reply(errorMessage);
        }
    }
});

client.login(process.env.DISCORD_TOKEN);