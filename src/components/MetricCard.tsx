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
    padding: 16,
    marginBottom: 12,
    borderRadius: 4,
    boxShadow: '0 2px 4px rgba(0, 0, 0, 0.1)',
    backgroundColor: 'white',
    borderLeft: `4px solid ${theme.palette.tealLight}`,
    transition: 'all 0.2s ease',
    cursor: 'pointer',
    ':hover': {
      boxShadow: '0 4px 8px rgba(0, 0, 0, 0.15)',
      transform: 'translateY(-2px)'
    }
  },
  selectedCard: {
    backgroundColor: theme.palette.neutralLighterAlt,
    boxShadow: '0 0 0 2px ' + theme.palette.themePrimary,
    borderLeft: `4px solid ${theme.palette.themeDarker}`,
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 12
  },
  title: {
    fontSize: 16,
    fontWeight: FontWeights.semibold,
    color: theme.palette.themePrimary
  },
  type: {
    display: 'flex',
    alignItems: 'center',
    fontSize: 12,
    padding: '3px 8px',
    borderRadius: 12,
    backgroundColor: theme.palette.neutralLighter,
    color: theme.palette.neutralPrimary,
    marginLeft: 8
  },
  icon: {
    fontSize: 14,
    marginRight: 6,
    color: theme.palette.neutralTertiary
  },
  detailsRow: {
    display: 'flex',
    fontSize: 13,
    marginBottom: 6,
    alignItems: 'center'
  },
  label: {
    width: 80,
    fontSize: 12,
    color: theme.palette.neutralSecondary
  },
  value: {
    flex: 1,
    color: theme.palette.neutralPrimary
  },
  description: {
    fontSize: 13,
    color: theme.palette.neutralSecondary,
    marginBottom: 12,
    maxHeight: 40,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    display: '-webkit-box',
    WebkitLineClamp: 2,
    WebkitBoxOrient: 'vertical'
  },
  footerSection: {
    borderTop: `1px solid ${theme.palette.neutralLighter}`,
    marginTop: 12,
    paddingTop: 8,
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    fontSize: 11,
    color: theme.palette.neutralSecondary
  },
  selectIndicator: {
    position: 'absolute',
    top: 10,
    right: 10,
    color: theme.palette.themePrimary,
    fontSize: 16
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
        year: 'numeric', 
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
        <Stack horizontal>
          <Text className={styles.title}>{item.label || item.name}</Text>
          <div className={styles.type}>
            <Icon iconName="Chart" className={styles.icon} />
            {item.assetType}
          </div>
        </Stack>
        <Icon iconName={getMetricIcon()} style={{ fontSize: 20, color: theme.palette.tealLight }} />
      </div>

      {item.description && (
        <div className={styles.description}>
          {item.description || "No description available"}
        </div>
      )}

      <div className={styles.detailsRow}>
        <span className={styles.label}>Created By:</span>
        <span className={styles.value}>
          <Icon iconName="Contact" style={{ fontSize: 14, marginRight: 6 }} />
          {item.createdBy?.Name || 'Unknown'}
        </span>
      </div>

      <div className={styles.detailsRow}>
        <span className={styles.label}>Created:</span>
        <span className={styles.value}>
          <Icon iconName="Calendar" style={{ fontSize: 14, marginRight: 6 }} />
          {formatDate(item.createdDate)}
        </span>
      </div>

      {item.semanticMetricId && (
        <div className={styles.detailsRow}>
          <span className={styles.label}>Semantic ID:</span>
          <span className={styles.value}>
            <Icon iconName="Link" style={{ fontSize: 14, marginRight: 6 }} />
            {item.semanticMetricId}
          </span>
        </div>
      )}

      <div className={styles.footerSection}>
        <span>
          <Icon iconName="ChartSeries" style={{ fontSize: 14, marginRight: 6 }} />
          Metric ID: {item.id}
        </span>
        <span>
          <Icon iconName="Clock" style={{ fontSize: 14, marginRight: 6 }} />
          {formatDate(item.lastModifiedDate)}
        </span>
      </div>
    </div>
  );
}; 