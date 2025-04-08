import React, { useState, useEffect } from 'react';
import { DefaultButton, Stack, Text, PrimaryButton, Spinner, SpinnerSize, MessageBar, MessageBarType, Image, ImageFit, Dialog, DialogType, Icon, getTheme, TextField, Pivot, PivotItem } from '@fluentui/react';
import { SalesforceAuth } from '../services/salesforce-real';
import { OfficeService } from '../services/office';
import { SalesforceDataCard } from './SalesforceDataCard';
import { MetricCard } from './MetricCard';

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

  const insertSelectedItemIntoDocument = async () => {
    if (!selectedItem) return;
    
    setLoading(true);
    setError(null);
    
    try {
      // Get the preview card element
      const previewElement = document.getElementById('preview-card');
      if (!previewElement) {
        throw new Error('Preview element not found');
      }

      // First insert the preview as an image
      await OfficeService.insertImageFromElement(previewElement);

      // Then insert the Lightning component
      if (selectedItem.Id) { // Check if it's a dashboard
        const accessToken = localStorage.getItem('sf_access_token');
        if (!accessToken) {
          throw new Error('Salesforce access token not found');
        }

        // Add some spacing between preview and component
        await OfficeService.insertHtml('<p style="margin: 20px 0;"></p>');

        // Insert the Lightning component
        await OfficeService.insertLightningComponent(
          'analytics_embedding:dashboard3p',
          {
            height: 300,
            idOrApiName: selectedItem.Id
          },
          accessToken
        );
      }

      setShowPreview(false);
      setError(null);
    } catch (err) {
      console.error('Error inserting content:', err);
      setError('Failed to insert content into document');
    } finally {
      setLoading(false);
    }
  };

  // This component renders a visual card for the selected item
  const PreviewCard = ({ item }: { item: any }) => {
    const theme = getTheme();
    
    const getDashboardIcon = () => {
      return 'ViewDashboard';
    };
    
    const getMetricIcon = () => {
      return 'BarChart4';
    };
    
    const formatDate = (dateString: string) => {
      if (!dateString) return 'Unknown';
      return new Date(dateString).toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      });
    };
    
    // Determine if this is a dashboard or metric
    const isMetric = item.hasOwnProperty('id') && item.hasOwnProperty('label');
    
    return (
      <Stack 
        id="preview-card"
        tokens={{ childrenGap: 12 }}
        styles={{
          root: {
            padding: 16,
            backgroundColor: theme.palette.white,
            borderRadius: 4,
            border: `1px solid ${theme.palette.neutralLight}`,
            boxShadow: theme.effects.elevation4
          }
        }}
      >
        {isMetric ? (
          // Metric Preview
          <>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName={getMetricIcon()} style={{ fontSize: 20, color: theme.palette.themePrimary }} />
                <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
                  {item.label || 'Unknown Metric'}
                </Text>
              </Stack>
              <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                {item.type || 'Metric'}
              </Text>
            </Stack>
            
            {item.description && (
              <Text variant="medium">
                {item.description}
              </Text>
            )}
            
            <Stack tokens={{ childrenGap: 8 }}>
              <Stack horizontal horizontalAlign="space-between">
                <Text variant="small">Creator:</Text>
                <Text variant="small">{item.creatorName || 'Unknown'}</Text>
              </Stack>
              
              <Stack horizontal horizontalAlign="space-between">
                <Text variant="small">Created:</Text>
                <Text variant="small">{formatDate(item.createdDate)}</Text>
              </Stack>
              
              <Stack horizontal horizontalAlign="space-between">
                <Text variant="small">Workspace:</Text>
                <Text variant="small">{item.namespace || 'Default'}</Text>
              </Stack>
            </Stack>
            
            <Stack 
              styles={{
                root: {
                  borderTop: `1px solid ${theme.palette.neutralLight}`,
                  paddingTop: 12
                }
              }}
            >
              <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                Metric ID: {item.id || 'Unknown'}
              </Text>
            </Stack>
          </>
        ) : (
          // Dashboard Preview
          <>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName={getDashboardIcon()} style={{ fontSize: 20, color: theme.palette.themePrimary }} />
                <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
                  {item.MasterLabel || item.Name || 'Unknown Dashboard'}
                </Text>
              </Stack>
              <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                {item.AnalyticsWorkspace ? item.AnalyticsWorkspace.MasterLabel : 'Dashboard'}
              </Text>
            </Stack>
            
            {item.Description && (
              <Text variant="medium">
                {item.Description}
              </Text>
            )}
            
            <Stack tokens={{ childrenGap: 8 }}>
              <Stack horizontal horizontalAlign="space-between">
                <Text variant="small">Creator:</Text>
                <Text variant="small">{item.CreatedBy ? item.CreatedBy.Name : 'Unknown'}</Text>
              </Stack>
              
              <Stack horizontal horizontalAlign="space-between">
                <Text variant="small">Created:</Text>
                <Text variant="small">{formatDate(item.CreatedDate)}</Text>
              </Stack>
              
              <Stack horizontal horizontalAlign="space-between">
                <Text variant="small">Workspace:</Text>
                <Text variant="small">{item.AnalyticsWorkspace ? item.AnalyticsWorkspace.MasterLabel : 'Default'}</Text>
              </Stack>
            </Stack>
            
            <Stack 
              styles={{
                root: {
                  borderTop: `1px solid ${theme.palette.neutralLight}`,
                  paddingTop: 12
                }
              }}
            >
              <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                Dashboard ID: {item.Id || 'Unknown'}
              </Text>
            </Stack>
          </>
        )}
      </Stack>
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
              type: DialogType.normal,
              title: selectedItem ? `Preview: ${selectedItem.MasterLabel || selectedItem.label || selectedItem.Name || selectedItem.name}` : 'Preview',
              subText: 'This is how the content will appear in your Office document'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { maxWidth: 650 } }
            }}
          >
            {selectedItem && (
              <Stack tokens={{ childrenGap: 15 }}>
                <Text variant="medium">
                  The following content would be inserted into your document:
                </Text>
                
                <Stack>
                  <PreviewCard item={selectedItem} />
                </Stack>
                
                <Stack horizontal tokens={{ childrenGap: 10 }} horizontalAlign="end">
                  <DefaultButton 
                    text="Cancel" 
                    onClick={() => setShowPreview(false)} 
                  />
                  <PrimaryButton 
                    text="Insert into Document" 
                    onClick={insertSelectedItemIntoDocument}
                    iconProps={{ iconName: 'Add' }}
                  />
                </Stack>
              </Stack>
            )}
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