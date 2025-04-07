import React from 'react';
import { 
  Stack, 
  Text, 
  FontWeights, 
  Icon, 
  mergeStyleSets, 
  getTheme,
} from '@fluentui/react';

interface MetricCardProps {
  item: any;
  onSelect?: (item: any) => void;
  isSelected?: boolean;
}

const theme = getTheme();
const styles = mergeStyleSets({
  card: {
    padding: '8px',
    marginBottom: '8px',
    borderRadius: 4,
    boxShadow: '0 1px 2px rgba(0, 0, 0, 0.1)',
    backgroundColor: 'white',
    borderLeft: `3px solid ${theme.palette.tealLight}`,
    transition: 'all 0.2s ease',
    cursor: 'pointer',
    ':hover': {
      boxShadow: '0 2px 4px rgba(0, 0, 0, 0.15)',
      transform: 'translateY(-1px)'
    }
  },
  selectedCard: {
    backgroundColor: theme.palette.neutralLighterAlt,
    boxShadow: '0 0 0 1px ' + theme.palette.themePrimary,
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 6
  },
  title: {
    fontSize: 14,
    fontWeight: FontWeights.semibold,
    color: theme.palette.themePrimary,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    maxWidth: '180px'
  },
  type: {
    display: 'flex',
    alignItems: 'center',
    fontSize: 11,
    padding: '2px 6px',
    borderRadius: 8,
    backgroundColor: theme.palette.neutralLighter,
    color: theme.palette.neutralPrimary,
    marginLeft: 6,
    maxWidth: '100px',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap'
  },
  icon: {
    fontSize: 12,
    marginRight: 4,
    color: theme.palette.neutralTertiary
  },
  detailsRow: {
    display: 'flex',
    fontSize: 11,
    marginBottom: 4,
    alignItems: 'center'
  },
  label: {
    width: 60,
    fontSize: 11,
    color: theme.palette.neutralSecondary
  },
  value: {
    flex: 1,
    color: theme.palette.neutralPrimary,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap'
  },
  description: {
    fontSize: 11,
    color: theme.palette.neutralSecondary,
    marginBottom: 6,
    maxHeight: 28,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    display: '-webkit-box',
    WebkitLineClamp: 2,
    WebkitBoxOrient: 'vertical'
  },
  footerSection: {
    borderTop: `1px solid ${theme.palette.neutralLighter}`,
    marginTop: 6,
    paddingTop: 4,
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    fontSize: 10,
    color: theme.palette.neutralSecondary
  },
  selectIndicator: {
    position: 'absolute',
    top: 6,
    right: 6,
    color: theme.palette.themePrimary,
    fontSize: 12
  }
});

export const MetricCard: React.FC<MetricCardProps> = ({ 
  item, 
  onSelect, 
  isSelected = false 
}) => {
  // Function to get an appropriate icon for the metric
  const getMetricIcon = () => {
    return 'BarChart4';
  };

  // Format the creation date
  const formatDate = (dateString: string) => {
    if (!dateString) return 'N/A';
    
    try {
      const date = new Date(dateString);
      return date.toLocaleDateString(undefined, { 
        year: '2-digit', 
        month: 'short', 
        day: 'numeric' 
      });
    } catch (e) {
      return dateString;
    }
  };

  const handleCardClick = () => {
    if (onSelect) {
      onSelect(item);
    }
  };

  return (
    <div 
      className={`${styles.card} ${isSelected ? styles.selectedCard : ''}`} 
      onClick={handleCardClick}
      role="button"
      tabIndex={0}
      aria-pressed={isSelected}
    >
      {isSelected && (
        <Icon iconName="CheckMark" className={styles.selectIndicator} />
      )}
      <div className={styles.header}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
          <Text className={styles.title} title={item.label || item.name}>
            {item.label || item.name}
          </Text>
          <div className={styles.type} title={item.assetType || 'Metric'}>
            <Icon iconName="Chart" className={styles.icon} />
            {item.assetType || 'Metric'}
          </div>
        </Stack>
      </div>

      {item.description && (
        <div className={styles.description} title={item.description}>
          {item.description || "No description available"}
        </div>
      )}

      <div className={styles.detailsRow}>
        <span className={styles.label}>Created:</span>
        <span className={styles.value}>
          <Icon iconName="Calendar" style={{ fontSize: 10, marginRight: 4 }} />
          {formatDate(item.createdDate)}
          <span style={{ margin: '0 4px', color: theme.palette.neutralTertiary }}>•</span>
          <Icon iconName="Contact" style={{ fontSize: 10, marginRight: 4 }} />
          {item.createdBy?.Name || 'Unknown'}
        </span>
      </div>

      <div className={styles.footerSection}>
        <span style={{ fontSize: 10 }}>
          <Icon iconName="ChartSeries" style={{ fontSize: 10, marginRight: 4 }} />
          ID: {item.id?.substring(0, 10)}...
        </span>
      </div>
    </div>
  );
}; 