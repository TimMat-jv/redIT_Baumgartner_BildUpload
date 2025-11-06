import React, { useEffect, useState, useMemo } from "react";
import { useMsal, useAccount } from "@azure/msal-react";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { loginRequest } from "../authConfig";
import ChannelsList from "./ChannelsList";
import { postMessageToChannel } from "./PostMessage";
import { Autocomplete, TextField, Button, Typography, Box, Alert, IconButton } from "@mui/material";
import { Star, StarBorder } from "@mui/icons-material";
import { db, OfflineDB, FavoriteTeam, OfflinePost, Team, Channel } from '../db';

const TeamsList: React.FC = () => {
    const { instance, accounts } = useMsal();
    const account = useAccount(accounts[0] || {});
    const [teams, setTeams] = useState<Team[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);
    const [selectedTeam, setSelectedTeam] = useState<Team | null>(null);
    const [selectedChannel, setSelectedChannel] = useState<Channel | null>(null);
    const [uploadSuccess, setUploadSuccess] = useState<boolean>(false);
    const [customText, setCustomText] = useState<string>("");
    const [imageUrls, setImageUrls] = useState<string[]>([]);
    const [posting, setPosting] = useState<boolean>(false);
    const [favorites, setFavorites] = useState<Set<string>>(new Set());
    const [isOnline, setIsOnline] = useState(navigator.onLine);
    const [offlinePosts, setOfflinePosts] = useState<any[]>([]);
    const [cachedFavorites, setCachedFavorites] = useState<any[]>([]);

    // Online-Status überwachen
    useEffect(() => {
        const handleOnline = () => setIsOnline(true);
        const handleOffline = () => setIsOnline(false);
        window.addEventListener('online', handleOnline);
        window.addEventListener('offline', handleOffline);
        return () => {
            window.removeEventListener('online', handleOnline);
            window.removeEventListener('offline', handleOffline);
        };
    }, []);

    // Lade gecachte Favoriten und Offline-Posts
    useEffect(() => {
        const loadCachedData = async () => {
            const cached = await db.favoriteTeams.toArray();
            setCachedFavorites(cached);
            const posts = await db.posts.toArray();
            setOfflinePosts(posts);
        };
        loadCachedData();
    }, []);

    useEffect(() => {
        const stored = localStorage.getItem('favoriteTeams');
        setFavorites(stored ? new Set(JSON.parse(stored)) : new Set());
    }, []);

    useEffect(() => {
        const fetchTeams = async () => {
            if (!account || !isOnline) return;  // Nur online laden

            const request = { ...loginRequest, account };

            try {
                const response = await instance.acquireTokenSilent(request);
                const accessToken = response.accessToken;

                const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me/joinedTeams", {
                    headers: { Authorization: `Bearer ${accessToken}` },
                });

                if (graphResponse.ok) {
                    const data = await graphResponse.json();
                    setTeams(data.value);
                } else {
                    setError("Failed to fetch teams");
                }
            } catch (err) {
                if (err instanceof InteractionRequiredAuthError) {
                    instance.acquireTokenPopup(request).then((response) => {
                        const accessToken = response.accessToken;
                        fetch("https://graph.microsoft.com/v1.0/me/joinedTeams", {
                            headers: { Authorization: `Bearer ${accessToken}` },
                        }).then((res) => res.json()).then((data) => setTeams(data.value));
                    });
                } else {
                    setError("Error fetching teams");
                }
            } finally {
                setLoading(false);
            }
        };

        fetchTeams();
    }, [instance, account, favorites, isOnline]);

    const toggleFavorite = async (teamId: string) => {
        const newFavorites = new Set(favorites);
        if (newFavorites.has(teamId)) {
            newFavorites.delete(teamId);
            await db.favoriteTeams.delete(teamId);  // Aus Cache entfernen
        } else {
            newFavorites.add(teamId);
            // Cache Team und Kanäle (nur online)
            if (isOnline && account) {
                const team = teams.find(t => t.id === teamId);
                if (team) {
                    // Kanäle laden und cachen
                    const request = { ...loginRequest, account };
                    const response = await instance.acquireTokenSilent(request);
                    const accessToken = response.accessToken;
                    const channelsResponse = await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/channels`, {
                        headers: { Authorization: `Bearer ${accessToken}` },
                    });
                    const channelsData = await channelsResponse.json();
                    await db.favoriteTeams.put({ id: teamId, displayName: team.displayName, channels: channelsData.value });
                }
            }
        }
        setFavorites(newFavorites);
        localStorage.setItem('favoriteTeams', JSON.stringify([...newFavorites]));
    };

    const handleTeamSelect = (event: any, value: Team | null) => {
        setSelectedTeam(value);
        setUploadSuccess(false);
        setCustomText("");
        setImageUrls([]);
    };

    const saveOfflinePost = async (files?: File[]) => {
        if (!selectedTeam || !selectedChannel) return;
        const post = {
            teamId: selectedTeam.id,
            channelId: selectedChannel.id,
            text: customText,
            imageUrls,
            timestamp: Date.now()
        };
        await db.posts.add(post);
        setOfflinePosts([...offlinePosts, post]);
        alert('Offline gespeichert!');
        setCustomText('');
        setImageUrls([]);
    };

    const syncOfflinePosts = async () => {
        if (!account || !isOnline || offlinePosts.length === 0) return;
        for (const post of offlinePosts) {
            try {
                const request = { ...loginRequest, account };
                const response = await instance.acquireTokenSilent(request);
                const accessToken = response.accessToken;
                await postMessageToChannel(accessToken, post.teamId, post.channelId, post.text, post.imageUrls);
                await db.posts.delete(post.id);
            } catch (err) {
                console.error('Sync-Fehler:', err);
            }
        }
        setOfflinePosts([]);
        alert('Offline-Posts synchronisiert!');
    };

    const handlePostToChannel = async () => {
        if (!account || !selectedTeam || !selectedChannel || !customText || imageUrls.length === 0) return;

        setPosting(true);

        const request = { ...loginRequest, account };

        try {
            const response = await instance.acquireTokenSilent(request);
            const accessToken = response.accessToken;

            await postMessageToChannel(accessToken, selectedTeam.id, selectedChannel.id, customText, imageUrls);

            alert("Beitrag erfolgreich in den Kanal gepostet!");
            setUploadSuccess(false);
            setCustomText("");
            setImageUrls([]);
        } catch (err) {
            alert("Fehler beim Posten: " + (err instanceof Error ? err.message : "Unbekannter Fehler"));
        } finally {
            setPosting(false);
        }
    };

    // Sortiere Teams: Favoriten zuerst
    const sortedTeams = useMemo(() => {
        return [...teams].sort((a, b) => {
            const aFav = favorites.has(a.id);
            const bFav = favorites.has(b.id);
            if (aFav && !bFav) return -1;
            if (!aFav && bFav) return 1;
            return a.displayName.localeCompare(b.displayName);
        });
    }, [teams, favorites]);

    if (loading) return <Typography variant="h6">Loading teams...</Typography>;
    if (error) return <Alert severity="error">Error: {error}</Alert>;

    return (
        <Box sx={{ mt: 3 }}>
            {/* Offline-Hinweis */}
            {(!isOnline || !account) && (
                <Alert severity="warning" sx={{ mb: 2 }}>
                    {!isOnline ? 'Offline-Modus: Eingaben werden lokal gespeichert.' : 'Nicht eingeloggt: Eingaben werden lokal gespeichert.'}
                </Alert>
            )}

            <Typography variant="h5" gutterBottom>
                Team auswählen
            </Typography>
            <Autocomplete
                options={sortedTeams}
                getOptionLabel={(option) => option.displayName}
                value={selectedTeam}
                onChange={handleTeamSelect}
                renderOption={(props, option) => (
                    <Box component="li" {...props} sx={{ display: 'flex', alignItems: 'center' }}>
                        <IconButton size="small" onClick={(e) => { e.stopPropagation(); toggleFavorite(option.id); }}>
                            {favorites.has(option.id) ? <Star color="primary" /> : <StarBorder />}
                        </IconButton>
                        {option.displayName}
                    </Box>
                )}
                renderInput={(params) => <TextField {...params} label="Search teams" variant="outlined" />}
                sx={{ mb: 2 }}
            />
            {selectedTeam && (
                <ChannelsList
                    team={selectedTeam}
                    onChannelSelect={setSelectedChannel}
                    onUploadSuccess={(urls: string[], files?: File[]) => {
                        setImageUrls(urls);
                        setUploadSuccess(true);
                        // Für Offline: Speichere files, wenn vorhanden
                        if (files && !isOnline) {
                            saveOfflinePost(files);
                        }
                    }}
                    onCustomTextChange={setCustomText}
                    customText={customText}
                    isFavorite={favorites.has(selectedTeam.id)}
                />
            )}
            {uploadSuccess && customText.trim() && (
                <Button
                    variant="contained"
                    color="primary"
                    onClick={isOnline && account ? handlePostToChannel : () => saveOfflinePost()}  // Wrappe saveOfflinePost in eine Funktion
                    disabled={posting}
                    sx={{ mt: 2 }}
                >
                    {posting ? "Posting..." : (isOnline && account ? "Beitrag in Kanal posten" : "Offline speichern")}
                </Button>
            )}
            {/* Sync-Button */}
            {isOnline && account && offlinePosts.length > 0 && (
                <Button onClick={syncOfflinePosts} variant="contained" sx={{ mt: 2 }}>
                    Offline-Posts synchronisieren ({offlinePosts.length})
                </Button>
            )}
        </Box>
    );
};

export default TeamsList;