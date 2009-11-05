/*
MySQL Data Transfer
Source Host: localhost
Source Database: ninjastory
Target Host: localhost
Target Database: ninjastory
Date: 2/23/2008 11:02:49 AM
*/

SET FOREIGN_KEY_CHECKS=0;
-- ----------------------------
-- Table structure for accounts
-- ----------------------------
CREATE TABLE `accounts` (
  `name` varchar(10) NOT NULL default '' COMMENT 'Account name',
  `pass` char(32) NOT NULL default '' COMMENT 'Account password''s MD5 hash',
  `user1` varchar(10) NOT NULL default '' COMMENT 'Name of character 1',
  `user2` varchar(10) NOT NULL default '' COMMENT 'Name of character 2',
  `user3` varchar(10) NOT NULL default '' COMMENT 'Name of character 3',
  `user4` varchar(10) NOT NULL default '' COMMENT 'Name of character 4',
  `user5` varchar(10) NOT NULL default '' COMMENT 'Name of character 5',
  PRIMARY KEY  (`name`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for items
-- ----------------------------
CREATE TABLE `items` (
  `id` smallint(6) NOT NULL COMMENT 'Unique ID of the item',
  `name` varchar(35) NOT NULL default '' COMMENT 'Name of the item (client-side only)',
  `desc` varchar(255) NOT NULL default '' COMMENT 'Description of the item (client-side only)',
  `type` enum('USEONCE','CAP','FOREHEAD','RING','EYEACC','EARACC','GLOVES','PENDANT','PANTS','CLOTHES','WEAPON','SHIELD','MANTLE','SHOES') NOT NULL default 'USEONCE' COMMENT 'Item type',
  `grhindex` int(11) NOT NULL default '0' COMMENT 'Grh index of the item (client-side only)',
  `stacking` smallint(6) NOT NULL default '1' COMMENT 'How much the item can stack in the inventory',
  `def` smallint(6) NOT NULL default '0' COMMENT 'Defense increased',
  `minhit` smallint(6) NOT NULL default '0' COMMENT 'Minimum hit damage increased',
  `maxhit` smallint(6) NOT NULL default '0' COMMENT 'Minimum hit damage increased',
  `hp` smallint(6) NOT NULL default '0' COMMENT 'HP replenished upon usage',
  `mp` smallint(6) NOT NULL default '0' COMMENT 'MP replenished upon usage',
  `maxhp` smallint(6) NOT NULL default '0' COMMENT 'Maximum HP increased',
  `maxmp` smallint(6) NOT NULL default '0' COMMENT 'Maximum MP increased',
  `str` smallint(6) NOT NULL default '0' COMMENT 'Strength increased',
  `dex` smallint(6) NOT NULL default '0' COMMENT 'Dexterity increased',
  `intl` smallint(6) NOT NULL default '0' COMMENT 'Intelligence increased',
  `luk` smallint(6) NOT NULL default '0' COMMENT 'Luck increased',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for npcs
-- ----------------------------
CREATE TABLE `npcs` (
  `id` smallint(6) NOT NULL COMMENT 'Unique ID for the NPC',
  `name` varchar(15) NOT NULL COMMENT 'The NPC''s name',
  `sprite` tinyint(3) unsigned NOT NULL default '1' COMMENT 'Index of the NPC''s sprite (\\Data\\Sprite.dat)',
  `spawn` smallint(6) NOT NULL default '5000' COMMENT 'Time (in ms) it takes for the NPC to spawn',
  `drop` text NOT NULL,
  `heading` enum('EAST','WEST') NOT NULL default 'EAST' COMMENT 'The NPC''s default heading',
  `stat_exp` smallint(6) NOT NULL default '0' COMMENT 'How much experience the NPC gives upon being killed',
  `stat_ryu` smallint(6) NOT NULL default '0' COMMENT 'How much Ryu the NPC gives upon being killed',
  `stat_hp` smallint(6) NOT NULL default '10' COMMENT 'Maximum health / start health',
  `stat_mp` smallint(6) NOT NULL default '10' COMMENT 'Maximum mana / start mana',
  `stat_str` smallint(6) NOT NULL default '1' COMMENT 'Base strength',
  `stat_dex` smallint(6) NOT NULL default '1' COMMENT 'Base dexterity',
  `stat_intl` smallint(6) NOT NULL default '1' COMMENT 'Base intelligence',
  `stat_luk` smallint(6) NOT NULL default '1' COMMENT 'Base luck',
  `stat_minhit` smallint(6) NOT NULL default '1' COMMENT 'Minimum hit damage',
  `stat_maxhit` smallint(6) NOT NULL default '1' COMMENT 'Maximum hit damage',
  `stat_def` smallint(6) NOT NULL default '0' COMMENT 'Defense',
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Table structure for users
-- ----------------------------
CREATE TABLE `users` (
  `name` varchar(10) NOT NULL default '' COMMENT 'Name of the character',
  `pos_x` smallint(6) NOT NULL default '10' COMMENT 'X co-ordinate',
  `pos_y` smallint(6) NOT NULL default '10' COMMENT 'Y co-ordinate',
  `pos_map` smallint(6) NOT NULL default '1' COMMENT 'Map index the user is on',
  `stat_lvl` smallint(6) NOT NULL default '1' COMMENT 'Level',
  `stat_exp` int(11) NOT NULL default '0' COMMENT 'Current experience points',
  `stat_ryu` int(11) NOT NULL default '0' COMMENT 'Current Ryu',
  `stat_str` smallint(6) NOT NULL default '1' COMMENT 'Base strength',
  `stat_dex` smallint(6) NOT NULL default '1' COMMENT 'Base dexterity',
  `stat_intl` smallint(6) NOT NULL default '1' COMMENT 'Base intelligence',
  `stat_luk` smallint(6) NOT NULL default '1' COMMENT 'Base luck',
  `stat_hp` smallint(6) NOT NULL default '10' COMMENT 'Current health',
  `stat_maxhp` smallint(6) NOT NULL default '10' COMMENT 'Maximum health',
  `stat_mp` smallint(6) NOT NULL default '10' COMMENT 'Current mana',
  `stat_maxmp` smallint(6) NOT NULL default '10' COMMENT 'Maximum mana',
  `eq_weapon` smallint(6) NOT NULL default '0' COMMENT 'Equipped weapon item index',
  `eq_cap` smallint(6) NOT NULL default '0' COMMENT 'Equipped cap item index',
  `eq_forehead` smallint(6) NOT NULL default '0' COMMENT 'Equipped forehead item index',
  `eq_ring1` smallint(6) NOT NULL default '0' COMMENT 'Equipped ring 1 item index',
  `eq_ring2` smallint(6) NOT NULL default '0' COMMENT 'Equipped ring 2 item index',
  `eq_ring3` smallint(6) NOT NULL default '0' COMMENT 'Equipped ring 3 item index',
  `eq_ring4` smallint(6) NOT NULL default '0' COMMENT 'Equipped ring 4 item index',
  `eq_eyeacc` smallint(6) NOT NULL default '0' COMMENT 'Equipped eye accessory item index',
  `eq_earacc` smallint(6) NOT NULL default '0' COMMENT 'Equipped ear accessory item index',
  `eq_gloves` smallint(6) NOT NULL default '0' COMMENT 'Equipped gloves item index',
  `eq_pendant` smallint(6) NOT NULL default '0' COMMENT 'Equipped pendant item index',
  `eq_pants` smallint(6) NOT NULL default '0' COMMENT 'Equipped pants item index',
  `eq_shoes` smallint(6) NOT NULL default '0' COMMENT 'Equipped shoes item index',
  `eq_shield` smallint(6) NOT NULL default '0' COMMENT 'Equipped shield item index',
  `eq_mantle` smallint(6) NOT NULL default '0' COMMENT 'Equipped mantle item index',
  `eq_clothes` smallint(6) NOT NULL default '0' COMMENT 'Equipped clothes item index',
  `char_body` tinyint(3) unsigned NOT NULL default '1' COMMENT 'Body paper-doll index',
  `inv` text NOT NULL COMMENT 'User''s inventory',
  PRIMARY KEY  (`name`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ----------------------------
-- Records 
-- ----------------------------
INSERT INTO `accounts` VALUES ('Spodi', '098f6bcd4621d373cade4e832627b4f6', 'Spodichr', '', '', '', '');
INSERT INTO `accounts` VALUES ('Test1', '098f6bcd4621d373cade4e832627b4f6', 'Test1', '', '', '', '');
INSERT INTO `accounts` VALUES ('Test10', '098f6bcd4621d373cade4e832627b4f6', 'Test10', '', '', '', '');
INSERT INTO `accounts` VALUES ('Test2', '098f6bcd4621d373cade4e832627b4f6', 'Test2', '', '', '', '');
INSERT INTO `accounts` VALUES ('Test3', '098f6bcd4621d373cade4e832627b4f6', 'Test3', '', '', '', '');
INSERT INTO `accounts` VALUES ('Test4', '098f6bcd4621d373cade4e832627b4f6', 'Test4', '', '', '', '');
INSERT INTO `accounts` VALUES ('Test5', '098f6bcd4621d373cade4e832627b4f6', 'Test5', '', '', '', '');
INSERT INTO `accounts` VALUES ('Test6', '098f6bcd4621d373cade4e832627b4f6', 'Test6', '', '', '', '');
INSERT INTO `accounts` VALUES ('Test7', '098f6bcd4621d373cade4e832627b4f6', 'Test7', '', '', '', '');
INSERT INTO `accounts` VALUES ('Test8', '098f6bcd4621d373cade4e832627b4f6', 'Test8', '', '', '', '');
INSERT INTO `accounts` VALUES ('Test9', '098f6bcd4621d373cade4e832627b4f6', 'Test9', '', '', '', '');
INSERT INTO `items` VALUES ('1', 'Potion', 'Delicious potion that restores your stupid life', 'USEONCE', '400', '100', '0', '0', '0', '10', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('2', 'Sword', 'Sword of pure pwnage.', 'WEAPON', '401', '1', '0', '100', '200', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('3', 'Cap', 'Just a cap', 'CAP', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('4', 'Bandana', 'A bandana', 'FOREHEAD', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('5', 'Ring one', 'A ring', 'RING', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('6', 'Ring two', 'A ring', 'RING', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('7', 'Ring three', 'A ring', 'RING', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('8', 'Ring four', 'A ring', 'RING', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('9', 'Glasses', 'Some glasses', 'EYEACC', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('10', 'Ear ring', 'A big fat loop ring', 'EARACC', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('11', 'Mittens', 'Keeps your hands warm', 'GLOVES', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('12', 'Necklace', 'It goes around your damn neck', 'PENDANT', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('13', 'Short shorts', 'Shows off that junk in your trunk', 'PANTS', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('14', 'Garbage bag', 'The ultimate body condom with maximum protection', 'CLOTHES', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('15', 'Garbage can lid', 'Block them bitches', 'SHIELD', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('16', 'A mantle', 'Who cares...', 'MANTLE', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `items` VALUES ('17', 'Nike shoes', 'Made by 4 year old Asian kids in a sweatshop', 'SHOES', '401', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0');
INSERT INTO `npcs` VALUES ('1', 'Evil Dude', '1', '5000', '1,1,100\r\n2,1,50', 'WEST', '5', '5', '10', '10', '1', '1', '1', '1', '1', '1', '0');
INSERT INTO `users` VALUES ('Spodichr', '100', '2009', '1', '1', '0', '0', '1', '1', '1', '1', '10', '10', '10', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '');
INSERT INTO `users` VALUES ('Test1', '435', '1817', '1', '1', '0', '0', '1', '1', '1', '1', '10', '10', '10', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '');
INSERT INTO `users` VALUES ('Test10', '10', '10', '1', '1', '0', '0', '1', '1', '1', '1', '10', '10', '10', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '');
INSERT INTO `users` VALUES ('Test2', '409', '1977', '1', '48', '25', '6135', '1', '1', '1', '1', '5', '10', '10', '10', '2', '3', '4', '8', '6', '7', '5', '9', '10', '11', '0', '13', '17', '15', '16', '14', '1', '1,57\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n2,1\r\n');
INSERT INTO `users` VALUES ('Test3', '863', '1305', '1', '1', '0', '0', '1', '1', '1', '1', '10', '10', '10', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '');
INSERT INTO `users` VALUES ('Test4', '10', '10', '1', '1', '0', '0', '1', '1', '1', '1', '10', '10', '10', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '');
INSERT INTO `users` VALUES ('Test5', '10', '10', '1', '1', '0', '0', '1', '1', '1', '1', '10', '10', '10', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '');
INSERT INTO `users` VALUES ('Test6', '10', '10', '1', '1', '0', '0', '1', '1', '1', '1', '10', '10', '10', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '');
INSERT INTO `users` VALUES ('Test7', '10', '10', '1', '1', '0', '0', '1', '1', '1', '1', '10', '10', '10', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '');
INSERT INTO `users` VALUES ('Test8', '10', '10', '1', '1', '0', '0', '1', '1', '1', '1', '10', '10', '10', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '');
INSERT INTO `users` VALUES ('Test9', '10', '10', '1', '1', '0', '0', '1', '1', '1', '1', '10', '10', '10', '10', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '');
