import React, { useState, useEffect, useRef } from 'react';
import { DefaultButton, Stack, Text, PrimaryButton, Spinner, SpinnerSize, MessageBar, MessageBarType, Image, ImageFit, Dialog, DialogType, Icon, getTheme, TextField, Pivot, PivotItem, DialogFooter } from '@fluentui/react';
import { SalesforceAuth } from '../services/salesforce-real';
import { OfficeService } from '../services/office';
import { SalesforceDataCard } from './SalesforceDataCard';
import { MetricCard } from './MetricCard';
import html2canvas from 'html2canvas';

interface AppProps {
  isLocalMode?: boolean;
}

export const App: React.FC<AppProps> = ({ isLocalMode = false }) => {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [dashboardData, setDashboardData] = useState<any[] | null>(null);
  const [metricsData, setMetricsData] = useState<any[] | null>(null);
  const [selectedItem, setSelectedItem] = useState<any | null>(null);
  const [showPreview, setShowPreview] = useState(false);
  const [showConfig, setShowConfig] = useState(false);
  const [clientId, setClientId] = useState(localStorage.getItem('sf_client_id') || '');
  const [activeTab, setActiveTab] = useState('dashboards');
  const [showClientInfo, setShowClientInfo] = useState(false);

  useEffect(() => {
    checkAuthStatus();
  }, []);

  const checkAuthStatus = async () => {
    try {
      setLoading(true);
      const auth = await SalesforceAuth.checkAuth();
      setIsAuthenticated(auth);
      
      if (auth) {
        // Automatically load data when authenticated
        loadAllData();
      }
    } catch (err) {
      console.error('Error checking authentication:', err);
      setError('Failed to check authentication status');
    } finally {
      setLoading(false);
    }
  };

  // Load both dashboards and metrics
  const loadAllData = async () => {
    setLoading(true);
    setError(null);
    
    try {
      if (isLocalMode) {
        // In local mode, fetch and display both types of data
        const dashboards = await SalesforceAuth.getData();
        setDashboardData(dashboards);
        console.log('Loaded dashboard data:', dashboards);
        
        const metrics = await SalesforceAuth.getMetricsData();
        setMetricsData(metrics);
        console.log('Loaded metrics data:', metrics);
      } else {
        // In Office mode, we still load both data types for selection
        // but don't insert anything automatically
        const dashboards = await SalesforceAuth.getData();
        setDashboardData(dashboards);
        
        const metrics = await SalesforceAuth.getMetricsData();
        setMetricsData(metrics);
        
        console.log('Office mode: Data loaded and ready for selection');
      }
    } catch (err) {
      console.error('Data retrieval error:', err);
      setError('Failed to retrieve Salesforce data');
    } finally {
      setLoading(false);
    }
  };

  const handleLogin = async () => {
    setLoading(true);
    setError(null);
    try {
      if (!localStorage.getItem('sf_client_id')) {
        setShowConfig(true);
        setLoading(false);
        return;
      }

      await SalesforceAuth.login();
      setIsAuthenticated(true);
      
      // Automatically load data after successful login
      await loadAllData();
    } catch (err) {
      console.error('Authentication error:', err);
      setError('Failed to authenticate with Salesforce');
    } finally {
      setLoading(false);
    }
  };

  const handleSaveConfig = () => {
    if (clientId.trim()) {
      localStorage.setItem('sf_client_id', clientId.trim());
      
      (SalesforceAuth as any).CLIENT_ID = clientId.trim();
      
      setShowConfig(false);
      setError(null);
      
      handleLogin();
    } else {
      setError('Client ID is required');
    }
  };

  const handleCardSelect = (item: any) => {
    setSelectedItem(item);
    setShowPreview(true);
    
    if (isLocalMode) {
      console.log('Selected item:', item);
    }
  };

  // Generate a placeholder image URL based on the company name and industry
  const getPlaceholderImageUrl = (item: any) => {
    const industry = item.Industry.toLowerCase().replace(/[^a-z0-9]/g, '');
    const companyName = item.Name.toLowerCase().replace(/[^a-z0-9]/g, '');
    // Using a different placeholder service that doesn't require encoding
    return `https://place-hold.it/600x300/0078D4/FFFFFF/bold?text=${companyName}`;
  };

  const handleInsertClick = async () => {
    if (!selectedItem) return;
    
    try {
      const previewElement = document.querySelector('.preview-card') as HTMLDivElement;
      if (!previewElement) {
        throw new Error('Preview element not found');
      }

      // Both metrics and dashboards now use Lightning Out
      console.log('Preparing to capture visualization...');
      
      const containerId = `preview-lightning-${selectedItem.DeveloperName || selectedItem.id || 'unknown'}`;
      const lightningContainer = document.getElementById(containerId);
      if (!lightningContainer) {
        throw new Error('Lightning container not found');
      }

      let attempts = 0;
      const maxAttempts = 30; // 30 seconds max wait time
      let visualContainer = null;
      
      // Wait for container within the Lightning container
      while (attempts < maxAttempts) {
        // For dashboards, look for dashboard container. For metrics, look for metric container
        visualContainer = selectedItem.hasOwnProperty('id') 
          ? lightningContainer.querySelector('.metric-container') 
          : lightningContainer.querySelector('.tua-dashboard-container');
          
        if (visualContainer) {
          console.log('Visual container found within Lightning Out, preparing to capture...');
          break;
        }
        console.log(`Waiting for content to load (attempt ${attempts + 1}/${maxAttempts})...`);
        await new Promise(resolve => setTimeout(resolve, 1000));
        attempts++;
      }

      if (!visualContainer) {
        console.error('Timed out waiting for content to load');
        throw new Error('Content failed to load');
      }

      // Give a small delay for final render
      await new Promise(resolve => setTimeout(resolve, 500));
      console.log('Capturing content...');

      // Capture just the visual container
      const canvas = await html2canvas(visualContainer as HTMLElement, {
        background: 'white',
        logging: false,
        useCORS: true,
        allowTaint: true
      });

      console.log('Content captured successfully');

      const imageData = {
        ...selectedItem,
        capturedImage: canvas.toDataURL('image/png')
      };

      await OfficeService.insertContent(imageData);
      console.log('Content inserted into document');

      setShowPreview(false);
    } catch (error) {
      console.error('Failed to capture and insert content:', error);
      setShowPreview(false);
    }
  };

  // This component renders a visual card for the selected item
  const PreviewCard = ({ item }: { item: any }) => {
    const theme = getTheme();
    const [isLightningLoaded, setIsLightningLoaded] = useState(false);
    const containerId = `preview-lightning-${item.DeveloperName || item.id || 'unknown'}`;
    const previewRef = useRef<HTMLDivElement>(null);
    
    useEffect(() => {
      let mounted = true;
      const cleanupTimeout: NodeJS.Timeout | null = null;
      
      // Initialize Lightning component for both metrics and dashboards
      const accessToken = localStorage.getItem('sf_access_token');
      if (accessToken) {
        // For metrics, we use the metric ID. For dashboards, we use the DeveloperName
        const idOrName = item.hasOwnProperty('id') ? item.id : item.DeveloperName;
        
        OfficeService.initializeLightningComponent(containerId, idOrName, accessToken)
          .then(() => {
            if (mounted) {
              setIsLightningLoaded(true);
            }
          })
          .catch(error => {
            console.error('Error initializing Lightning component:', error);
            if (mounted) {
              setIsLightningLoaded(false);
            }
          });
      }
      
      return () => {
        mounted = false;
        if (cleanupTimeout) clearTimeout(cleanupTimeout);
        
        // Cleanup Lightning component
        try {
          const container = document.getElementById(containerId);
          if (container) {
            container.remove();
          }
        } catch (error) {
          console.warn('Error during Lightning component cleanup:', error);
        }
      };
    }, [item.DeveloperName, item.id, containerId]);
    
    const formatDate = (dateString: string) => {
      if (!dateString) return 'Unknown';
      return new Date(dateString).toLocaleDateString('en-US', {
        year: '2-digit',
        month: 'short',
        day: 'numeric'
      });
    };
    
    // Determine if this is a metric or dashboard
    const isMetric = item.hasOwnProperty('id') && item.hasOwnProperty('metadata');

    // Common preview layout for both metrics and dashboards
    return (
      <div ref={previewRef} className="preview-card" style={{ padding: '20px' }}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Icon 
                iconName={isMetric ? "Chart" : "ViewDashboard"} 
                style={{ fontSize: 20, color: theme.palette.themePrimary }} 
              />
              <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
                {isMetric 
                  ? (item.metadata?.asset?.label || item.label)
                  : (item.MasterLabel || item.Name || 'Unknown Dashboard')
                }
              </Text>
            </Stack>
            <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
              {isMetric ? 'Metric' : (item.AnalyticsWorkspace ? item.AnalyticsWorkspace.MasterLabel : 'Dashboard')}
            </Text>
          </Stack>

          {/* Lightning Component Container */}
          <div style={{ 
            minHeight: 300, 
            border: `1px solid ${theme.palette.neutralLight}`,
            borderRadius: 4,
            padding: 8,
            backgroundColor: theme.palette.white,
            position: 'relative'
          }}>
            <div id={containerId}>
              {!isLightningLoaded && (
                <Stack 
                  horizontalAlign="center" 
                  verticalAlign="center" 
                  styles={{ root: { height: 300 } }}
                >
                  <Spinner size={SpinnerSize.large} label={`Loading ${isMetric ? 'metric' : 'dashboard'}...`} />
                </Stack>
              )}
            </div>
          </div>

          {/* Additional details for metrics */}
          {isMetric && item.metadata?.asset?.metricInsight && (
            <div style={{ 
              padding: '12px', 
              backgroundColor: theme.palette.neutralLighter,
              borderRadius: '4px'
            }}>
              <Text>
                <Icon iconName="Lightbulb" style={{ marginRight: 8, color: theme.palette.themePrimary }} />
                {item.metadata.asset.metricInsight}
              </Text>
            </div>
          )}

          <Stack 
            styles={{
              root: {
                borderTop: `1px solid ${theme.palette.neutralLight}`,
                paddingTop: 12,
                marginTop: 12
              }
            }}
          >
            <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
              {isMetric ? `Metric ID: ${item.metadata?.asset?.id || item.id}` : `Dashboard ID: ${item.Id || 'Unknown'}`}
            </Text>
          </Stack>
        </Stack>
      </div>
    );
  };

  return (
    <Stack tokens={{ childrenGap: 8, padding: '12px 8px' }}>
      <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
        <Text variant="large" styles={{ root: { fontWeight: 600 } }}>Tableau Next + Office</Text>
        
        {isAuthenticated && (
          <div style={{ 
            display: 'flex', 
            alignItems: 'center',
            position: 'relative'
          }}>
            <div 
              style={{ 
                backgroundColor: '#107C10', 
                color: 'white', 
                padding: '2px 8px',
                borderRadius: '12px',
                fontSize: '11px',
                fontWeight: 600,
                display: 'flex',
                alignItems: 'center',
                cursor: 'pointer'
              }}
              onMouseEnter={() => setShowClientInfo(true)}
              onMouseLeave={() => setShowClientInfo(false)}
            >
              <Icon iconName="PlugConnected" style={{ marginRight: 4, fontSize: '10px' }} />
              Connected
            </div>
            
            {showClientInfo && (
              <div style={{
                position: 'absolute',
                top: '100%',
                right: 0,
                marginTop: '6px',
                backgroundColor: 'white',
                boxShadow: '0 2px 8px rgba(0, 0, 0, 0.15)',
                padding: '6px 10px',
                borderRadius: '4px',
                zIndex: 100,
                fontSize: '11px',
                whiteSpace: 'nowrap'
              }}>
                Using Salesforce Client ID: {localStorage.getItem('sf_client_id')?.substring(0, 8)}...
              </div>
            )}
          </div>
        )}
      </Stack>
      
      {error && (
        <MessageBar 
          messageBarType={MessageBarType.error} 
          onDismiss={() => setError(null)}
          styles={{
            root: { padding: '8px' },
            text: { fontSize: '11px' }
          }}
        >
          {error}
        </MessageBar>
      )}

      {loading && (
        <Spinner 
          size={SpinnerSize.small}
          label={isAuthenticated ? "Loading Salesforce data..." : "Processing authentication..."} 
          styles={{ 
            root: { margin: '6px 0' },
            label: { fontSize: '11px' }
          }}
        />
      )}

      {!isAuthenticated ? (
        <Stack tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text="Connect to Salesforce"
            onClick={handleLogin}
            disabled={loading}
            iconProps={{ iconName: 'Plug' }}
            styles={{
              root: { height: '32px' },
              label: { fontSize: '12px' }
            }}
          />
          
          <DefaultButton
            text="Configure Salesforce Connection"
            onClick={() => setShowConfig(true)}
            iconProps={{ iconName: 'Settings' }}
            styles={{
              root: { height: '32px' },
              label: { fontSize: '12px' }
            }}
          />
        </Stack>
      ) : (
        <Stack tokens={{ childrenGap: 8 }}>
          {(dashboardData || metricsData) && (
            <Stack tokens={{ childrenGap: 4 }}>
              <Pivot 
                selectedKey={activeTab} 
                onLinkClick={(item) => item && setActiveTab(item.props.itemKey || 'dashboards')}
                styles={{ 
                  root: { marginBottom: 8 },
                  link: { height: '32px', minWidth: 'auto' },
                  linkContent: { fontSize: '12px' },
                  count: { fontSize: '11px', marginLeft: '4px' }
                }}
              >
                <PivotItem 
                  headerText="Dashboards" 
                  itemKey="dashboards" 
                  itemCount={dashboardData?.length || 0}
                  headerButtonProps={{
                    'data-order': 1,
                    'data-title': 'Dashboards'
                  }}
                />
                <PivotItem 
                  headerText="Metrics" 
                  itemKey="metrics" 
                  itemCount={metricsData?.length || 0}
                  headerButtonProps={{
                    'data-order': 2,
                    'data-title': 'Metrics'
                  }}
                />
              </Pivot>
              
              <div style={{ maxHeight: 'calc(100vh - 180px)', overflowY: 'auto', padding: '0 4px', marginRight: '-4px' }}>
                {activeTab === 'dashboards' && dashboardData && dashboardData.map((item, index) => (
                  <SalesforceDataCard 
                    key={item.Id || index} 
                    item={item} 
                    onSelect={handleCardSelect}
                    isSelected={selectedItem && selectedItem.Id === item.Id}
                  />
                ))}
                
                {activeTab === 'metrics' && metricsData && metricsData.map((item, index) => (
                  <MetricCard 
                    key={item.id || index} 
                    item={item} 
                    onSelect={handleCardSelect}
                    isSelected={selectedItem && selectedItem.id === item.id}
                  />
                ))}
                
                {activeTab === 'dashboards' && (!dashboardData || dashboardData.length === 0) && (
                  <MessageBar styles={{ text: { fontSize: '11px' } }}>
                    No dashboard data available.
                  </MessageBar>
                )}
                
                {activeTab === 'metrics' && (!metricsData || metricsData.length === 0) && (
                  <MessageBar styles={{ text: { fontSize: '11px' } }}>
                    No metrics data available.
                  </MessageBar>
                )}
              </div>
            </Stack>
          )}
          
          {/* Preview Dialog for selected item */}
          <Dialog
            hidden={!showPreview}
            onDismiss={() => setShowPreview(false)}
            dialogContentProps={{
              title: 'Metric Preview'
            }}
            modalProps={{
              styles: { main: { maxWidth: 700 } }
            }}
          >
            {selectedItem && <PreviewCard item={selectedItem} />}
            <DialogFooter>
              <PrimaryButton onClick={handleInsertClick} text="Insert into Document" />
              <DefaultButton onClick={() => setShowPreview(false)} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </Stack>
      )}

      {/* Salesforce Configuration Dialog */}
      <Dialog
        hidden={!showConfig}
        onDismiss={() => setShowConfig(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Salesforce Configuration',
          subText: 'Enter your Salesforce OAuth client ID to connect to your Salesforce instance'
        }}
      >
        <Stack tokens={{ childrenGap: 15 }}>
          <TextField
            label="Salesforce Client ID"
            value={clientId}
            onChange={(_e, newValue) => setClientId(newValue || '')}
            required
            placeholder="Enter your Salesforce client ID"
          />
          
          <Text variant="small">
            To obtain a client ID:
            <ol style={{ margin: '8px 0', paddingLeft: 20 }}>
              <li>Go to Salesforce Setup</li>
              <li>Navigate to App Manager</li>
              <li>Create a New Connected App</li>
              <li>Enable OAuth settings and add callback URL: <code style={{ background: '#f0f0f0', padding: '2px 4px' }}>{window.location.origin}/oauth-callback.html</code></li>
              <li>Make sure the callback URL in your Connected App matches exactly</li>
              <li>Copy the Consumer Key (Client ID)</li>
            </ol>
          </Text>
          
          <Stack horizontal tokens={{ childrenGap: 10 }} horizontalAlign="end">
            <DefaultButton text="Cancel" onClick={() => setShowConfig(false)} />
            <PrimaryButton text="Save" onClick={handleSaveConfig} />
          </Stack>
        </Stack>
      </Dialog>
    </Stack>
  );
}; 